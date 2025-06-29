import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
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
    id: zod.alias('i', z.string())
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    title: zod.alias('t', z.string().optional()),
    url: zod.alias('u', z.string().optional())
      .refine(url => url !== undefined && (validation.isValidSharePointUrl(url)), {
        message: `The value for 'url' must be a valid SharePoint URL or a server-relative URL starting with '/'`
      })
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
    return 'Retrieves information of a specific site as admin';
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
      await logger.logToStderr(`Retrieving information about site ${args.options.id || args.options.title || args.options.url}...`);
    }

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.verbose);
      let tenantSiteInfo: TenantSiteProperties | undefined;

      if (args.options.id) {
        tenantSiteInfo = await this.getTenantSiteById(spoAdminUrl, args.options.id);
      }
      else if (args.options.title) {
        const allSites: TenantSiteProperties[] = await spo.getAllSites(spoAdminUrl, logger, this.verbose, '', false, '');
        const filterSites: TenantSiteProperties[] = allSites.filter(site => site.Title?.toLowerCase() === args.options.title?.toLowerCase());

        if (filterSites.length === 0) {
          throw `No site found with title '${args.options.title}'`;
        }
        else if (filterSites.length === 1) {
          tenantSiteInfo = await this.getTenantSiteById(spoAdminUrl, formatting.extractCsomGuid(filterSites[0].SiteId!));
        }
        else {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('Url', filterSites);
          const result = await cli.handleMultipleResultsFound<{ SiteId: string }>(`Multiple sites with title '${args.options.title}' found.`, resultAsKeyValuePair);

          tenantSiteInfo = await this.getTenantSiteById(spoAdminUrl, formatting.extractCsomGuid(result.SiteId));
        }
      }
      else if (args.options.url) {
        tenantSiteInfo = await spo.getSiteAdminPropertiesByUrl(args.options.url, false, logger, this.verbose);
      }

      await logger.log(tenantSiteInfo);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTenantSiteById(spoAdminUrl: string, id: string): Promise<TenantSiteProperties> {
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/sites('${id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<TenantSiteProperties>(requestOptions);
  }
}

export default new SpoTenantSiteGetCommand();