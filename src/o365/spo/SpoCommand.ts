import Command, { CommandAction, CommandError } from '../../Command';
import auth from './SpoAuth';
import * as request from 'request-promise-native';
import Utils from '../../Utils';
import { SpoOperation } from './commands/site/SpoOperation';
import config from '../../config';

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
          cmd.initAction(args);

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

  protected getRequestDigest(cmd: CommandInstance, debug: boolean): request.RequestPromise {
    return this.getRequestDigestForSite(auth.site.url, auth.site.accessToken, cmd, debug);
  }

  protected getRequestDigestForSite(siteUrl: string, accessToken: string, cmd: CommandInstance, debug: boolean): request.RequestPromise {
    const requestOptions: any = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.post(requestOptions);
  }

  public static isValidSharePointUrl(url: string): boolean | string {
    if (!url) {
      return false;
    }

    if (url.indexOf('https://') !== 0 ||
      url.indexOf('.sharepoint.com') === -1) {
      return `${url} is not a valid SharePoint Online site URL`;
    }
    else {
      return true;
    }
  }
}