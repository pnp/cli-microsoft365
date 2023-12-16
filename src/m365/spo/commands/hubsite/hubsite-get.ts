import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemListCommand, { Options as SpoListItemListCommandOptions } from '../listitem/listitem-list.js';
import { AssociatedSite } from './AssociatedSite.js';
import { HubSite } from './HubSite.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  url?: string;
  includeAssociatedSites?: boolean;
}

class SpoHubSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified hub site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        url: typeof args.options.url !== 'undefined',
        includeAssociatedSites: args.options.includeAssociatedSites === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '-u, --url [url]' },
      { option: '--includeAssociatedSites' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.url) {
          return validation.isValidSharePointUrl(args.options.url);
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title', 'url'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const hubSite = args.options.id ? await this.getHubSiteById(spoUrl, args.options) : await this.getHubSite(spoUrl, args.options);

      if (args.options.includeAssociatedSites && (args.options.output && args.options.output !== 'json')) {
        throw 'includeAssociatedSites option is only allowed with json output mode';
      }

      if (args.options.includeAssociatedSites === true && args.options.output && !cli.shouldTrimOutput(args.options.output)) {
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
