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
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const appCatalogs = await odata.getAllItems<any>(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);

      if (appCatalogs && appCatalogs.length > 0) {
        logger.log(appCatalogs);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No site collection app catalogs found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoSiteAppCatalogListCommand();