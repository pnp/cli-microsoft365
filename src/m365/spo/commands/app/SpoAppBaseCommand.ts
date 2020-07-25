import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  scope?: string;
}

export abstract class SpoAppBaseCommand extends SpoCommand {
  public getAppCatalogSiteUrl(cmd: CommandInstance, authSiteUrl: string, args: CommandArgs): Promise<string> {
    return new Promise<string>((resolve: (appCatalogSiteUrl: string) => void, reject: (error: any) => void): void => {
      if (args.options.scope === 'sitecollection') {
        return resolve((args.options.appCatalogUrl as string).toLowerCase().replace('/appcatalog', ''));
      }

      if (args.options.appCatalogUrl) {
        return resolve(args.options.appCatalogUrl);
      }

      const requestOptions: any = {
        url: `${authSiteUrl}/_api/SP_TenantSettings_Current`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .get<{ CorporateCatalogUrl?: string }>(requestOptions)
        .then((res: { CorporateCatalogUrl?: string }) => {
          if (res.CorporateCatalogUrl) {
            return resolve(res.CorporateCatalogUrl);
          }

          reject('Tenant app catalog is not configured.');
        })
        .catch((err: any) => {
          reject(err);
        });
    });
  }
}