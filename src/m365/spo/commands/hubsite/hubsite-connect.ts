import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { HubSite } from './HubSite.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).optional().alias('u'),
  parentId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID`
  }).optional(),
  parentTitle: z.string().optional(),
  parentUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteConnectCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_CONNECT;
  }

  public get description(): string {
    return 'Connects a hub site to a parent hub site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.url].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id', 'title', or 'url', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'url']
        }
      })
      .refine(opts => [opts.parentId, opts.parentTitle, opts.parentUrl].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'parentId', 'parentTitle', or 'parentUrl', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['parentId', 'parentTitle', 'parentUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Connecting hub site '${args.options.id || args.options.title || args.options.url}' to hub site '${args.options.parentId || args.options.parentTitle || args.options.parentUrl}'...`);
    }

    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const hubSites = await this.getHubSites(spoAdminUrl);

      const hubSite = await this.getSpecificHubSite(hubSites, args.options.id, args.options.title, args.options.url);
      const parentHubSite = await this.getSpecificHubSite(hubSites, args.options.parentId, args.options.parentTitle, args.options.parentUrl);

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/HubSites/GetById('${hubSite.ID}')`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'if-match': hubSite['odata.etag']!
        },
        responseType: 'json',
        data: {
          ParentHubSiteId: parentHubSite.ID
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getHubSites(spoAdminUrl: string): Promise<HubSite[]> {
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/HubSites?$select=ID,Title,SiteUrl&$top=5000`,
      headers: {
        accept: 'application/json;odata=minimalmetadata'
      },
      responseType: 'json'
    };

    const hubSites = await request.get<{ value: HubSite[] }>(requestOptions);
    return hubSites.value;
  }

  private async getSpecificHubSite(hubSites: HubSite[], id?: string, title?: string, url?: string): Promise<HubSite> {
    let filteredHubSites: HubSite[] = [];

    if (id) {
      filteredHubSites = hubSites.filter(site => site.ID.toLowerCase() === id.toLowerCase());
    }
    else if (title) {
      filteredHubSites = hubSites.filter(site => site.Title.toLowerCase() === title.toLowerCase());
    }
    else if (url) {
      filteredHubSites = hubSites.filter(site => site.SiteUrl.toLowerCase() === url.toLowerCase());
    }

    if (filteredHubSites.length === 0) {
      throw `The specified hub site '${id || title || url}' does not exist.`;
    }
    if (filteredHubSites.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', filteredHubSites);
      return await cli.handleMultipleResultsFound<HubSite>(`Multiple hub sites with name '${title}' found.`, resultAsKeyValuePair);
    }

    return filteredHubSites[0];
  }
}

export default new SpoHubSiteConnectCommand();