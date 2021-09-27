import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
import config from '../../config';
import request from '../../request';
import { SpoOperation } from '../spo/commands/site/SpoOperation';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, FormDigestInfo } from '../spo/spo';
const csomDefs = require('../../../csom.json');

export interface FormDigest {
  formDigestValue: string;
  formDigestExpiresAt: Date;
}

export default abstract class SpoCommand extends Command {
  /**
   * Defines list of options that contain URLs in spo commands. CLI will use
   * this list to expand server-relative URLs specified in these options to
   * absolute.
   * If a command requires one of these options to contain a server-relative
   * URL, it should override this method and remove the necessary property from
   * the array before returning it.
   */
  protected getNamesOfOptionsWithUrls(): string[] {
    const namesOfOptionsWithUrls: string[] = [
      'appCatalogUrl',
      'siteUrl',
      'webUrl',
      'origin',
      'url',
      'imageUrl',
      'actionUrl',
      'logoUrl',
      'libraryUrl',
      'thumbnailUrl',
      'targetUrl',
      'newSiteUrl',
      'previewImageUrl',
      'NoAccessRedirectUrl',
      'StartASiteFormUrl',
      'OrgNewsSiteUrl',
      'parentWebUrl',
      'siteLogoUrl'
    ];
    const excludedOptionsWithUrls: string[] | undefined = this.getExcludedOptionsWithUrls();
    if (!excludedOptionsWithUrls) {
      return namesOfOptionsWithUrls;
    }
    else {
      return namesOfOptionsWithUrls.filter(o => excludedOptionsWithUrls.indexOf(o) < 0);
    }
  }

  /**
   * Array of names of options with URLs that should be excluded
   * from processing. To be overriden in commands that require
   * specific options to be a server-relative URL
   */
  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return undefined;
  }

  protected getRequestDigest(siteUrl: string): Promise<FormDigestInfo> {
    const requestOptions: any = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  public async processOptions(options: any): Promise<void> {
    const namesOfOptionsWithUrls: string[] = this.getNamesOfOptionsWithUrls();
    const optionNames = Object.getOwnPropertyNames(options);
    for (const optionName of optionNames) {
      if (namesOfOptionsWithUrls.indexOf(optionName) < 0) {
        continue;
      }

      const optionValue: any = options[optionName];
      if (typeof optionValue !== 'string' ||
        !optionValue.startsWith('/')) {
        continue;
      }

      await auth.restoreAuth();

      if (!auth.service.spoUrl) {
        throw new Error(`SharePoint URL is not available. Set SharePoint URL using the 'm365 spo set' command or use absolute URLs`);
      }

      options[optionName] = auth.service.spoUrl + optionValue;
    }
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

  public ensureFormDigest(siteUrl: string, logger: Logger, context: FormDigestInfo | undefined, debug: boolean): Promise<FormDigestInfo> {
    return new Promise<FormDigestInfo>((resolve: (context: FormDigestInfo) => void, reject: (error: any) => void): void => {
      if (this.isValidFormDigest(context)) {
        if (debug) {
          logger.logToStderr('Existing form digest still valid');
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

  protected waitUntilFinished(operationId: string, siteUrl: string, resolve: () => void, reject: (error: any) => void, logger: Logger, currentContext: FormDigestInfo, dots?: string): void {
    this
      .ensureFormDigest(siteUrl, logger, currentContext, this.debug)
      .then((res: FormDigestInfo): Promise<string> => {
        currentContext = res;

        if (this.debug) {
          logger.logToStderr(`Checking if operation ${operationId} completed...`);
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
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
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
          const isComplete: boolean = operation.IsComplete;
          if (isComplete) {
            if (!this.debug && this.verbose) {
              process.stdout.write('\n');
            }

            resolve();
            return;
          }

          setTimeout(() => {
            this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), siteUrl, resolve, reject, logger, currentContext, dots);
          }, operation.PollingInterval);
        }
      });
  }

  protected waitUntilCopyJobFinished(copyJobInfo: any, siteUrl: string, pollingInterval: number, resolve: () => void, reject: (error: any) => void, logger: Logger, dots?: string): void {
    const requestUrl: string = `${siteUrl}/_api/site/GetCopyJobProgress`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: { "copyJobInfo": copyJobInfo },
      responseType: 'json'
    };

    if (!this.debug && this.verbose) {
      dots += '.';
      process.stdout.write(`\r${dots}`);
    }

    request
      .post<{ JobState?: number, Logs: string[] }>(requestOptions)
      .then((resp: { JobState?: number, Logs: string[] }): void => {

        if (this.debug) {
          logger.logToStderr('getCopyJobProgress response...');
          logger.logToStderr(resp);
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
        }
        else {
          setTimeout(() => {
            this.waitUntilCopyJobFinished(copyJobInfo, siteUrl, pollingInterval, resolve, reject, logger, dots);
          }, pollingInterval);
        }
      });
  }

  protected getSpoUrl(logger: Logger, debug: boolean): Promise<string> {
    if (auth.service.spoUrl) {
      if (debug) {
        logger.logToStderr(`SPO URL previously retrieved ${auth.service.spoUrl}. Returning...`);
      }

      return Promise.resolve(auth.service.spoUrl);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`No SPO URL available. Retrieving from MS Graph...`);
      }

      const requestOptions: any = {
        url: `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
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

  protected getSpoAdminUrl(logger: Logger, debug: boolean): Promise<string> {
    return new Promise<string>((resolve: (spoAdminUrl: string) => void, reject: (error: any) => void): void => {
      this
        .getSpoUrl(logger, debug)
        .then((spoUrl: string): void => {
          resolve(spoUrl.replace(/(https:\/\/)([^\.]+)(.*)/, '$1$2-admin$3'));
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  protected getTenantId(logger: Logger, debug: boolean): Promise<string> {
    if (auth.service.tenantId) {
      if (debug) {
        logger.logToStderr(`SPO Tenant ID previously retrieved ${auth.service.tenantId}. Returning...`);
      }

      return Promise.resolve(auth.service.tenantId);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`No SPO Tenant ID available. Retrieving...`);
      }

      let spoAdminUrl: string = '';

      this
        .getSpoAdminUrl(logger, debug)
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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        if (auth.service.connected && AuthType[auth.service.authType] === AuthType[AuthType.Secret]) {
          cb(new CommandError(`SharePoint does not support authentication using client ID and secret. Please use a different login type to use SharePoint commands.`));
          return;
        }

        super.action(logger, args, cb);
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }
}
