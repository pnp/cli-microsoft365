import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  appCatalogScope?: string;
}

export abstract class SpoAppBaseCommand extends SpoCommand {
  protected async getAppCatalogSiteUrl(logger: Logger, authSiteUrl: string, args: CommandArgs): Promise<string> {
    if (args.options.appCatalogScope === 'sitecollection') {
      // Trim any trailing slashes if there are any
      const appCatalogUrl = args.options.appCatalogUrl!.replace(/\/$/, '');
      const appCatalogUrlChunks = appCatalogUrl.split('/');

      // Trim the last part of the URL if it ends on '/appcatalog', but don't trim it if the site URL is called like that (/sites/appcatalog).
      if (appCatalogUrl.toLowerCase().endsWith('/appcatalog') && appCatalogUrlChunks.length !== 5) {
        return appCatalogUrl.substring(0, appCatalogUrl.lastIndexOf('/'));
      }
    }

    if (args.options.appCatalogUrl) {
      return args.options.appCatalogUrl!.replace(/\/$/, '');
    }

    if (this.verbose) {
      logger.logToStderr('Getting tenant app catalog url...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${authSiteUrl}/_api/SP_TenantSettings_Current`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ CorporateCatalogUrl?: string }>(requestOptions);
    if (response.CorporateCatalogUrl) {
      return response.CorporateCatalogUrl;
    }

    throw new Error('Tenant app catalog is not configured.');
  }
}