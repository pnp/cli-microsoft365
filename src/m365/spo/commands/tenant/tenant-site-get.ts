import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { TenantSiteProperties } from './TenantSiteProperties.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    id: z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional(),
    title: z.string().optional(),
    url: z.string().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SITE_GET;
  }

  public get description(): string {
    return 'Retrieves the tenant site information';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.id, options.title, options.url].filter(x => x !== undefined).length === 1, {
        message: `Specify either id, title, or url, but not multiple.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about the tenant site information...`);
    }

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);

      if (args.options.id) {
        const tenantSite = await this.getTenantSiteById(spoAdminUrl, args.options.id);
        await logger.log(tenantSite);
      }
      else if (args.options.title) {
        const allSites: TenantSiteProperties[] = await spo.getAllSites(spoAdminUrl, '', '0', '0', '', undefined, logger, this.verbose);

        // Filter this.allSites by title
        const filterSites: TenantSiteProperties[] = allSites.filter(site => site.Title === args.options.title);

        if (filterSites.length === 0) {
          throw new Error(`No sites found with title '${args.options.title}'`);
        }
        else if (filterSites.length === 1) {
          const tenantSite = await this.getTenantSiteById(spoAdminUrl, this.replaceString(filterSites[0].SiteId!));
          await logger.log(tenantSite);
        }
        else {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('SiteId', filterSites);
          const result = await cli.handleMultipleResultsFound<{ SiteId: string }>(`Multiple sites with title '${args.options.title}' found.`, resultAsKeyValuePair);

          const tenantSite = await this.getTenantSiteById(spoAdminUrl, this.replaceString(result.SiteId));
          await logger.log(tenantSite);
        }
      }
      else if (args.options.url) {
        const tenantSite = await spo.getSiteAdminPropertiesByUrl(args.options.url, false, logger, this.verbose);
        await logger.log(tenantSite);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private replaceString(s: string): string {
    return s.replace('/Guid(', '').replace(')/', '');
  }

  private async getTenantSiteById(spoAdminUrl: string, id: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/sites('${id}')`,
      headers: {
        'content-type': 'application/json;charset=utf-8',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<any>(requestOptions);
  }
}

export default new SpoTenantSiteGetCommand();