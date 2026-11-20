import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { HubSite } from './HubSite.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).optional().alias('u'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteDisconnectCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects a hub site from its parent hub site';
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const disconnectHubSite = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Disconnecting hub site '${args.options.id || args.options.title || args.options.url}' from its parent hub site...`);
        }

        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const hubSite = await this.getHubSite(spoAdminUrl, args.options);

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_api/HubSites/GetById('${hubSite.ID}')`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'if-match': hubSite['odata.etag']
          },
          responseType: 'json',
          data: {
            ParentHubSiteId: '00000000-0000-0000-0000-000000000000'
          }
        };

        await request.patch(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await disconnectHubSite();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want disconnect hub site '${args.options.id || args.options.title || args.options.url}' from its parent hub site?` });

      if (result) {
        await disconnectHubSite();
      }
    }
  }

  private async getHubSite(spoAdminUrl: string, options: Options): Promise<{ 'odata.etag': string, ID: string }> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata=minimalmetadata'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${spoAdminUrl}/_api/HubSites/GetById('${options.id}')?$select=ID`;

      const result = await request.get<{ 'odata.etag': string, ID: string }>(requestOptions);
      return result;
    }

    requestOptions.url = `${spoAdminUrl}/_api/HubSites?$select=ID,Title,SiteUrl&$top=5000`;
    const hubSitesResponse = await request.get<{ value: HubSite[] }>(requestOptions);
    const hubSites = hubSitesResponse.value;

    let filteredHubSites: HubSite[] = [];
    if (options.title) {
      filteredHubSites = hubSites.filter(site => site.Title.toLowerCase() === options.title!.toLowerCase());
    }
    else if (options.url) {
      filteredHubSites = hubSites.filter(site => site.SiteUrl.toLowerCase() === options.url!.toLowerCase());
    }

    if (filteredHubSites.length === 0) {
      throw `The specified hub site '${options.title || options.url}' does not exist.`;
    }

    if (filteredHubSites.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', filteredHubSites);
      const result = await cli.handleMultipleResultsFound<HubSite>(`Multiple hub sites with name '${options.title}' found.`, resultAsKeyValuePair);
      return result as { 'odata.etag': string, ID: string };
    }

    return filteredHubSites[0] as { 'odata.etag': string, ID: string };
  }
}

export default new SpoHubSiteDisconnectCommand();