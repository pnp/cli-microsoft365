import Command, { CommandAction, CommandError } from '../../Command';
import appInsights from '../../appInsights';
import auth from './SpoAuth';
import { SearchResponse } from './spo';
import * as request from 'request-promise-native';
import Utils from '../../Utils';

export default abstract class SpoCommand extends Command {
  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public action(): CommandAction {
    const cmd: SpoCommand = this;

    return function (this: CommandInstance, args: any, cb: () => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd._debug = args.options.debug || false;
          cmd._verbose = cmd._debug || args.options.verbose || false;

          appInsights.trackEvent({
            name: cmd.getCommandName(),
            properties: cmd.getTelemetryProperties(args)
          });
          appInsights.flush();

          if (!auth.site.connected) {
            this.log(new CommandError('Connect to a SharePoint Online site first'));
            cb();
            return;
          }

          if (cmd.requiresTenantAdmin()) {
            if (!auth.site.isTenantAdminSite()) {
              this.log(new CommandError(`${auth.site.url} is not a tenant admin site. Connect to your tenant admin site and try again`));
              cb();
              return;
            }
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          this.log(new CommandError(error));
          cb();
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

  protected getTenantAppCatalogUrl(cmd: CommandInstance, debug: boolean): Promise<string> {
    return new Promise<string>((resolve: (appCatalogUrl: string) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${auth.site.url}/_api/search/query?querytext='contentclass:STS_Site%20AND%20SiteTemplate:APPCATALOG'&SelectProperties='SPWebUrl'`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${auth.site.accessToken}`,
          accept: 'application/json;odata=nometadata'
        }),
        json: true
      };

      if (debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      request
        .get(requestOptions)
        .then((res: SearchResponse): void => {
          if (debug) {
            cmd.log('Response');
            cmd.log(res);
            cmd.log('');
          }

          if (res.PrimaryQueryResult.RelevantResults.RowCount < 1) {
            reject('Tenant app catalog not found');
            return;
          }

          for (let i: number = 0; i < res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.length; i++) {
            if (res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells[i].Key === 'SPWebUrl') {
              resolve(res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells[i].Value);
              return;
            }
          }

          reject('Tenant app catalog URL not found');
        }, (error: any): void => {
          if (debug) {
            cmd.log('Error');
            cmd.log(error);
            cmd.log('');
          }

          reject(error);
        });
    });
  }
}