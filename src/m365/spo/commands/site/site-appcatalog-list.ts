import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  excludeDeletedSites: z.boolean().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAppCatalogListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_APPCATALOG_LIST;
  }

  public get description(): string {
    return 'List all site collection app catalogs within the tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['AbsoluteUrl', 'SiteID'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving site collection app catalogs...');
      }

      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      let appCatalogs = await odata.getAllItems<any>(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);

      if (args.options.excludeDeletedSites) {
        if (this.verbose) {
          await logger.logToStderr('Excluding inaccessible sites from the results...');
        }

        const activeAppCatalogs = [];
        for (const appCatalog of appCatalogs) {
          try {
            await spo.getWeb(appCatalog.AbsoluteUrl, logger, this.verbose);
            activeAppCatalogs.push(appCatalog);
          }
          catch (error: any) {
            if (this.debug) {
              await logger.logToStderr(error);
            }

            if (error.status === 404 || error.status === 403) {
              if (this.verbose) {
                await logger.logToStderr(`Site at '${appCatalog.AbsoluteUrl}' is inaccessible. Excluding from results...`);
              }
              continue;
            }

            throw error;
          }
        }

        appCatalogs = activeAppCatalogs;
      }

      await logger.log(appCatalogs);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppCatalogListCommand();