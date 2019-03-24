import Command, { CommandAction, CommandError } from '../../Command';
import auth from './SpoAuth';
import request from '../../request';
import { SpoOperation } from './commands/site/SpoOperation';
import config from '../../config';
import { FormDigestInfo, ClientSvcResponse, ClientSvcResponseContents } from './spo';

export interface FormDigest {
  formDigestValue: string; 
  formDigestExpiresAt: Date; 
}

export default abstract class SpoCommand extends Command {
  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public action(): CommandAction {
    const cmd: SpoCommand = this;

    return function (this: CommandInstance, args: any, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd.initAction(args, this);

          if (!auth.site.connected) {
            cb(new CommandError('Log in to a SharePoint Online site first'));
            return;
          }

          if (cmd.requiresTenantAdmin()) {
            if (!auth.site.isTenantAdminSite()) {
              cb(new CommandError(`${auth.site.url} is not a tenant admin site. Log in to your tenant admin site and try again`));
              return;
            }
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }

  protected getRequestDigest(cmd: CommandInstance, debug: boolean): Promise<FormDigestInfo> {
    return this.getRequestDigestForSite(auth.site.url, auth.site.accessToken, cmd, debug);
  }

  protected getRequestDigestForSite(siteUrl: string, accessToken: string, cmd: CommandInstance, debug: boolean): Promise<FormDigestInfo> {
    const requestOptions: any = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        authorization: `Bearer ${accessToken}`,
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

  public ensureFormDigest(cmd: CommandInstance, context: FormDigestInfo | undefined, debug: boolean): Promise<FormDigestInfo> {
    return new Promise<FormDigestInfo>((resolve: (context: FormDigestInfo) => void, reject: (error: any) => void): void => {
      if (this.isValidFormDigest(context)) {
        if (debug) {
          cmd.log('Existing form digest still valid');
        }

        resolve(context as FormDigestInfo);
        return; 
      }

      this
        .getRequestDigest(cmd, debug)
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

  protected waitUntilFinished(operationId: string, resolve: () => void, reject: (error: any) => void, accessToken: string, cmd: CommandInstance, currentContext: FormDigestInfo, dots?: string, timeout?: NodeJS.Timer): void {
    this
      .ensureFormDigest(cmd, currentContext, this.debug)
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
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
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
            this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), resolve, reject, accessToken, cmd, currentContext, dots);
          }, operation.PollingInterval);
        }
      });
  }
}