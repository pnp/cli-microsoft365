import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { HubSite } from './HubSite';
import { AssociatedSite } from './AssociatedSite';
import { Options as SpoListItemListCommandOptions } from '../listitem/listitem-list';
import * as SpoListItemListCommand from '../listitem/listitem-list';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  includeAssociatedSites?: boolean;
  title?: string;
  url?: string;
}

class SpoHubSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified hub site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.url = typeof args.options.url !== 'undefined';
    telemetryProps.includeAssociatedSites = args.options.includeAssociatedSites === true;
    return telemetryProps;
  }

  private getAssociatedSites(spoAdminUrl: string, hubSiteId: string, logger: Logger, args: CommandArgs): Promise<CommandOutput> {
    const options: SpoListItemListCommandOptions = {
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose,
      listTitle: 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS',
      webUrl: spoAdminUrl,
      filter: `HubSiteId eq '${hubSiteId}'`,
      fields: 'Title,SiteUrl,SiteId'
    };

    return Cli
      .executeCommandWithOutput(SpoListItemListCommand as Command, { options: { ...options, _: [] } });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let hubSite: HubSite;

    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<any> => {
        if (args.options.id) {
          return this.getHubSiteById(spoUrl, args.options);
        }
        else {
          return this.getHubSite(spoUrl, args.options);
        }
      })
      .then((res: HubSite): Promise<CommandOutput | void> => {
        hubSite = res;

        if (args.options.includeAssociatedSites && (args.options.output && args.options.output !== 'json')) {
          throw Error('includeAssociatedSites option is only allowed with json output mode');
        }

        if (args.options.includeAssociatedSites !== true || args.options.output && args.options.output !== 'json') {
          return Promise.resolve();
        }

        return spo
          .getSpoAdminUrl(logger, this.debug)
          .then((spoAdminUrl: string): Promise<CommandOutput> => {
            return this.getAssociatedSites(spoAdminUrl, hubSite.SiteId, logger, args);
          });
      })
      .then((associatedSitesCommandOutput: CommandOutput | void): void => {
        if (args.options.includeAssociatedSites) {
          const associatedSites: AssociatedSite[] = JSON.parse((associatedSitesCommandOutput as CommandOutput).stdout) as AssociatedSite[];
          hubSite.AssociatedSites = associatedSites.filter(s => s.SiteId !== hubSite.SiteId);
        }

        logger.log(hubSite);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getHubSiteById(spoUrl: string, options: Options): Promise<HubSite> {
    const requestOptions: any = {
      url: `${spoUrl}/_api/hubsites/getbyid('${options.id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.get(requestOptions);
  }

  private getHubSite(spoUrl: string, options: Options): Promise<HubSite> {
    const requestOptions: any = {
      url: `${spoUrl}/_api/hubsites`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => {
        let hubSites = response.value as HubSite[];

        if (options.title) {
          hubSites = hubSites.filter(site => site.Title.toLocaleLowerCase() === options.title!.toLocaleLowerCase());
        }
        else if (options.url) {
          hubSites = hubSites.filter(site => site.SiteUrl.toLocaleLowerCase() === options.url!.toLocaleLowerCase());
        }

        if (hubSites.length === 0) {
          return Promise.reject(`The specified hub site ${options.title || options.url} does not exist`);
        }

        if (hubSites.length > 1) {
          return Promise.reject(`Multiple hub sites with ${options.title || options.url} found. Please disambiguate: ${hubSites.map(site => site.SiteUrl).join(', ')}`);
        }

        return hubSites[0];
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-i, --id [id]' },
      { option: '-t, --title [title]' },
      { option: '-u, --url [url]' },
      { option: '--includeAssociatedSites' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public optionSets(): string[][] | undefined {
    return [
      ['id', 'title', 'url']
    ];
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id && !validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.url) {
      return validation.isValidSharePointUrl(args.options.url);
    }

    return true;
  }
}

module.exports = new SpoHubSiteGetCommand();
