import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  appCatalogScope?: string;
}

export abstract class SpoAppBaseCommand extends SpoCommand {
  public getAppCatalogSiteUrl(logger: Logger, authSiteUrl: string, args: CommandArgs): Promise<string> {
    return new Promise<string>((resolve: (appCatalogSiteUrl: string) => void, reject: (error: any) => void): void => {
      if (args.options.appCatalogScope === 'sitecollection') {
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
        responseType: 'json'
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