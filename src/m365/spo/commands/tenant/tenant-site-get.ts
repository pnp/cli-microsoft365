import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';

const optionsSchema = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), { message: 'Specify a valid GUID' }).optional()),
    title: zod.alias('t', z.string().optional()),
    url: zod.alias('u', z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
      message: 'Specify a valid SharePoint site URL'
    }).optional())
  })
  .strict();
declare type Options = z.infer<typeof optionsSchema>;

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

  public get schema(): z.ZodTypeAny | undefined {
    return optionsSchema;
  }

  public getRefinedSchema(schema: typeof optionsSchema): z.ZodEffects<any> | undefined {
    return schema.refine(o => [o.id, o.title, o.url].filter(v => v !== undefined).length === 1, {
      message: `Specify exactly one of the following options: 'id', 'title', or 'url'.`
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving tenant site information for site '${args.options.url || args.options.id || args.options.title}'...`);
    }

    try {
      let siteUrl: string;
      if (args.options.url) {
        siteUrl = args.options.url;
      }
      else if (args.options.id) {
        siteUrl = await this.getSiteUrlById(args.options.id, logger);
        if (this.verbose) {
          await logger.logToStderr(`Retrieved tenant site URL for site '${args.options.id}'...`);
        }
      }
      else {
        siteUrl = await this.getSiteUrlByTitle(args.options.title!, logger);
        if (this.verbose) {
          await logger.logToStderr(`Retrieved tenant site URL for site '${args.options.title}'...`);
        }
      }
      const site = await spo.getSiteAdminPropertiesByUrl(siteUrl, false, logger, this.verbose);
      await logger.log(site);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSiteUrlById(id: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving tenant site URL for site '${id}'...`);
    }

    const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/sites('${id}')?$select=Url`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<{ Url: string }>(requestOptions);
    return res.Url;
  }

  private async getSiteUrlByTitle(title: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving tenant site URL for site '${title}'...`);
    }

    const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
    const viewXml = `<View><Query><Where><And><IsNull><FieldRef Name="TimeDeleted"/></IsNull><Eq><FieldRef Name="Title"/><Value Type='Text'>${formatting.escapeXml(title)}</Value></Eq></And></Where></Query><ViewFields><FieldRef Name="Title"/><FieldRef Name="SiteUrl"/><FieldRef Name="SiteId"/></ViewFields></View>`;
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        parameters: {
          ViewXml: viewXml,
          DatesInUtc: true
        }
      }
    };

    const res = await request.post<{ Row: { Title: string; SiteUrl: string; SiteId: string }[] }>(requestOptions);
    const rows = res.Row;
    if (rows.length === 0) {
      throw `The specified site '${title}' does not exist.`;
    }

    if (rows.length > 1) {
      const resultAsKeyValuePair = rows.reduce((acc, cur) => {
        acc[cur.SiteUrl] = { url: cur.SiteUrl };
        return acc;
      }, {} as any);
      const selection = await cli.handleMultipleResultsFound<{ url: string }>(`Multiple sites with title '${title}' found.`, resultAsKeyValuePair);
      return selection.url;
    }

    return rows[0].SiteUrl;
  }
}

export default new SpoTenantSiteGetCommand();


