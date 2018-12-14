import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  scope?: string;
  siteUrl?: string;
}

export abstract class SpoAppBaseCommand extends SpoCommand {

  public getAppCatalogSiteUrl(cmd: CommandInstance, authSiteUrl: string, accessToken: string, args: CommandArgs): Promise<string> {

    return new Promise<string>((resolve: any, reject: any): void => {
     
      if (args.options.scope === 'sitecollection') {
        return resolve(args.options.siteUrl as string);
      }

      if (args.options.appCatalogUrl) {
        return resolve(args.options.appCatalogUrl);
      }

      const requestOptions: any = {
        url: `${authSiteUrl}/_api/SP_TenantSettings_Current`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${accessToken}`,
          accept: 'application/json;odata=nometadata'
        })
      };
  
      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }
  
      request.get(requestOptions)
      .then((res: any) => {

        if (this.debug) {
          cmd.log('Tenant App catalog response...');
          cmd.log(res);
          cmd.log('');
        }

        const json = JSON.parse(res);
        if(json.CorporateCatalogUrl) {
          return resolve(json.CorporateCatalogUrl as string);
        }
        return reject('Tenant app catalog is not configured.');
      })
      .catch((err: any) => {
        reject(err);
      });
    });
  }
}