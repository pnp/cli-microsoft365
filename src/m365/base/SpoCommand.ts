import Command from '../../Command';
import auth, { Logger } from '../../Auth';
import request from '../../request';
import { SpoOperation } from '../spo/commands/site/SpoOperation';
import config from '../../config';
import { FormDigestInfo, ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../spo/spo';
import { CommandInstance } from '../../cli';
const csomDefs = require('../../../csom.json');

export interface FormDigest {
  formDigestValue: string;
  formDigestExpiresAt: Date;
}

export default abstract class SpoCommand extends Command {
  protected getRequestDigest(siteUrl: string): Promise<FormDigestInfo> {
    const requestOptions: any = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    return request.post(requestOptions);
  }

  public static isValidSharePointUrl(url: string): boolean | string {
    if (!url) {
      return false;
    }

    if (url.indexOf('https://') !== 0) {
      return `${url} is not a valid SharePoint Online site URL`;
    }
    else {
      return true;
    }
  }

  public ensureFormDigest(siteUrl: string, cmd: CommandInstance, context: FormDigestInfo | undefined, debug: boolean): Promise<FormDigestInfo> {
    return new Promise<FormDigestInfo>((resolve: (context: FormDigestInfo) => void, reject: (error: any) => void): void => {
      if (this.isValidFormDigest(context)) {
        if (debug) {
          cmd.log('Existing form digest still valid');
        }

        resolve(context as FormDigestInfo);
        return;
      }

      this
        .getRequestDigest(siteUrl)
        .then((res: FormDigestInfo): void => {
          const now: Date = new Date();
          now.setSeconds(now.getSeconds() + res.FormDigestTimeoutSeconds - 5);
          context = {
            FormDigestValue: res.FormDigestValue,
            FormDigestTimeoutSeconds: res.FormDigestTimeoutSeconds,
            FormDigestExpiresAt: now,
            WebFullUrl: res.WebFullUrl
          };

          resolve(context);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private isValidFormDigest(contextInfo: FormDigestInfo | undefined): boolean {
    if (!contextInfo) {
      return false;
    }

    const now: Date = new Date();
    if (contextInfo.FormDigestValue && now < contextInfo.FormDigestExpiresAt) {
      return true;
    }

    return false;
  }

  protected waitUntilFinished(operationId: string, siteUrl: string, resolve: () => void, reject: (error: any) => void, cmd: CommandInstance, currentContext: FormDigestInfo, dots?: string, timeout?: NodeJS.Timer): void {
    this
      .ensureFormDigest(siteUrl, cmd, currentContext, this.debug)
      .then((res: FormDigestInfo): Promise<string> => {
        currentContext = res;

        if (this.debug) {
          cmd.log(`Checking if operation ${operationId} completed...`);
        }

        if (!this.debug && this.verbose) {
          dots += '.';
          process.stdout.write(`\r${dots}`);
        }

        const requestOptions: any = {
          url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': currentContext.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          reject(response.ErrorInfo.ErrorMessage);
        }
        else {
          const operation: SpoOperation = json[json.length - 1];
          let isComplete: boolean = operation.IsComplete;
          if (isComplete) {
            if (this.verbose) {
              process.stdout.write('\n');
            }

            resolve();
            return;
          }

          timeout = setTimeout(() => {
            this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), siteUrl, resolve, reject, cmd, currentContext, dots);
          }, operation.PollingInterval);
        }
      });
  }

  protected waitUntilCopyJobFinished(copyJobInfo: any, siteUrl: string, pollingInterval: number, resolve: () => void, reject: (error: any) => void, cmd: CommandInstance, dots?: string, timeout?: NodeJS.Timer): void {
    const requestUrl: string = `${siteUrl}/_api/site/GetCopyJobProgress`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      body: { "copyJobInfo": copyJobInfo },
      json: true
    };

    if (!this.debug && this.verbose) {
      dots += '.';
      process.stdout.write(`\r${dots}`);
    }

    request
      .post<{ JobState?: number, Logs: string[] }>(requestOptions)
      .then((resp: { JobState?: number, Logs: string[] }): void => {

        if (this.debug) {
          cmd.log('getCopyJobProgress response...');
          cmd.log(resp);
        }

        for (const item of resp.Logs) {
          const log: { Event: string; Message: string } = JSON.parse(item);

          // reject if progress error
          if (log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // two possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        if (resp.JobState === 0) {
          // job done
          if (this.verbose) {
            process.stdout.write('\n');
          }

          resolve();
        } else {
          timeout = setTimeout(() => {
            this.waitUntilCopyJobFinished(copyJobInfo, siteUrl, pollingInterval, resolve, reject, cmd, dots);
          }, pollingInterval);
        }
      });
  }

  protected getSpoUrl(stdout: Logger, debug: boolean): Promise<string> {
    if (auth.service.spoUrl) {
      if (debug) {
        stdout.log(`SPO URL previously retrieved ${auth.service.spoUrl}. Returning...`);
      }

      return Promise.resolve(auth.service.spoUrl);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        stdout.log(`No SPO URL available. Retrieving from MS Graph...`);
      }

      const requestOptions: any = {
        url: `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .get<{ webUrl: string }>(requestOptions)
        .then((res: { webUrl: string }): Promise<void> => {
          auth.service.spoUrl = res.webUrl;
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          resolve(auth.service.spoUrl as string);
        }, (err: any): void => {
          if (auth.service.spoUrl) {
            resolve(auth.service.spoUrl);
          }
          else {
            reject(err);
          }
        });
    });
  }

  protected getSpoAdminUrl(stdout: Logger, debug: boolean): Promise<string> {
    return new Promise<string>((resolve: (spoAdminUrl: string) => void, reject: (error: any) => void): void => {
      this
        .getSpoUrl(stdout, debug)
        .then((spoUrl: string): void => {
          resolve(spoUrl.replace(/(https:\/\/)([^\.]+)(.*)/, '$1$2-admin$3'));
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  protected getTenantId(stdout: Logger, debug: boolean): Promise<string> {
    if (auth.service.tenantId) {
      if (debug) {
        stdout.log(`SPO Tenant ID previously retrieved ${auth.service.tenantId}. Returning...`);
      }

      return Promise.resolve(auth.service.tenantId);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        stdout.log(`No SPO Tenant ID available. Retrieving...`);
      }

      let spoAdminUrl: string = '';

      this
        .getSpoAdminUrl(stdout, debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        })
        .then((contextInfo: ContextInfo): Promise<string> => {
          const tenantInfoRequestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': contextInfo.FormDigestValue,
              accept: 'application/json;odata=nometadata'
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
          };

          return request.post(tenantInfoRequestOptions);
        })
        .then((res: string): Promise<void> => {
          const json: string[] = JSON.parse(res);
          auth.service.tenantId = (json[json.length - 1] as any)._ObjectIdentity_.replace('\n', '&#xA;');
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          resolve(auth.service.tenantId as string);
        }, (err: any): void => {
          if (auth.service.tenantId) {
            resolve(auth.service.tenantId);
          }
          else {
            reject(err);
          }
        });
    });
  }

  protected validateUnknownOptions(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    const unknownOptions: any = this.getUnknownOptions(options);
    const optionNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    if (optionNames.length === 0) {
      return true;
    }

    for (let i: number = 0; i < optionNames.length; i++) {
      const optionName: string = optionNames[i];
      const csomOptionType: string = csomDefs[csomObject][csomPropertyType][optionName];

      if (!csomOptionType) {
        return `${optionName} is not a valid ${csomObject} property`;
      }

      if (['Boolean', 'String', 'Int32'].indexOf(csomOptionType) < 0) {
        return `Unknown properties of type ${csomOptionType} are not yet supported`;
      }
    }

    return true;
  }

  /**
   * Combines base and relative url considering any missing slashes
   * @param baseUrl https://contoso.com
   * @param relativeUrl sites/abc
   */
  protected urlCombine(baseUrl: string, relativeUrl: string): string {
    // remove last '/' of base if exists
    if (baseUrl.lastIndexOf('/') === baseUrl.length - 1) {
      baseUrl = baseUrl.substring(0, baseUrl.length - 1);
    }

    // remove '/' at 0
    if (relativeUrl.charAt(0) === '/') {
      relativeUrl = relativeUrl.substring(1, relativeUrl.length);
    }

    // remove last '/' of next if exists
    if (relativeUrl.lastIndexOf('/') === relativeUrl.length - 1) {
      relativeUrl = relativeUrl.substring(0, relativeUrl.length - 1);
    }

    return `${baseUrl}/${relativeUrl}`;
  }
}
