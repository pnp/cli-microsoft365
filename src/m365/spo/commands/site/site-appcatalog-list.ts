import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoSiteAppCatalogListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_APPCATALOG_LIST;
  }

  public get description(): string {
    return 'List all site collection app catalogs within the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['AbsoluteUrl', 'SiteID'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving site collection app catalogs...');
      }

      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const appCatalogs = await odata.getAllItems<any>(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);
      await logger.log(appCatalogs);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppCatalogListCommand();