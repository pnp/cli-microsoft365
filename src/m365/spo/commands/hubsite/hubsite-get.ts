import { z } from 'zod';
import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemListCommand, { Options as SpoListItemListCommandOptions } from '../listitem/listitem-list.js';
import { AssociatedSite } from './AssociatedSite.js';
import { HubSite } from './HubSite.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).optional().alias('u'),
  withAssociatedSites: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified hub site';
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
    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const hubSite = args.options.id ? await this.getHubSiteById(spoUrl, args.options) : await this.getHubSite(spoUrl, args.options);

      if (args.options.withAssociatedSites && (args.options.output && args.options.output !== 'json')) {
        throw 'withAssociatedSites option is only allowed with json output mode';
      }

      if (args.options.withAssociatedSites === true && args.options.output && !cli.shouldTrimOutput(args.options.output)) {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const associatedSitesCommandOutput = await this.getAssociatedSites(spoAdminUrl, hubSite.SiteId, logger, args);
        const associatedSites: AssociatedSite[] = JSON.parse((associatedSitesCommandOutput as CommandOutput).stdout) as AssociatedSite[];
        hubSite.AssociatedSites = associatedSites.filter(s => s.SiteId !== hubSite.SiteId);
      }

      await logger.log(hubSite);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAssociatedSites(spoAdminUrl: string, hubSiteId: string, logger: Logger, args: CommandArgs): Promise<CommandOutput> {
    const options: SpoListItemListCommandOptions = {
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose,
      listTitle: 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS',
      webUrl: spoAdminUrl,
      filter: `HubSiteId eq '${hubSiteId}'`,
      fields: 'Title,SiteUrl,SiteId'
    };

    return cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...options, _: [] } });
  }

  private async getHubSiteById(spoUrl: string, options: Options): Promise<HubSite> {
    const requestOptions: CliRequestOptions = {
      url: `${spoUrl}/_api/hubsites/getbyid('${options.id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.get(requestOptions);
  }

  private async getHubSite(spoUrl: string, options: Options): Promise<HubSite> {
    const requestOptions: CliRequestOptions = {
      url: `${spoUrl}/_api/hubsites`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: HubSite[] }>(requestOptions);
    let hubSites = response.value as HubSite[];

    if (options.title) {
      hubSites = hubSites.filter(site => site.Title.toLocaleLowerCase() === options.title!.toLocaleLowerCase());
    }
    else if (options.url) {
      hubSites = hubSites.filter(site => site.SiteUrl.toLocaleLowerCase() === options.url!.toLocaleLowerCase());
    }

    if (hubSites.length === 0) {
      throw `The specified hub site ${options.title || options.url} does not exist`;
    }

    if (hubSites.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', hubSites);
      return await cli.handleMultipleResultsFound<HubSite>(`Multiple hub sites with ${options.title || options.url} found.`, resultAsKeyValuePair);
    }

    return hubSites[0];
  }
}

export default new SpoHubSiteGetCommand();
