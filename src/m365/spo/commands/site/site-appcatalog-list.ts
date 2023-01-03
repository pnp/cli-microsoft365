import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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
        logger.logToStderr('Retrieving site collection app catalogs...');
      }

      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const appCatalogs = await odata.getAllItems<any>(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);
      logger.log(appCatalogs);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppCatalogListCommand();