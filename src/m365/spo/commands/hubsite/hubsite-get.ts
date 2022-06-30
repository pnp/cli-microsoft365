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
  id: string;
  includeAssociatedSites?: boolean;
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
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.includeAssociatedSites = args.options.includeAssociatedSites === true;
    return telemetryProps;
  }

  private getAssociatedSites(spoAdminUrl: string, logger: Logger, args: CommandArgs): Promise<CommandOutput> {
    const options: SpoListItemListCommandOptions = {
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose,
      listTitle: 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS',
      webUrl: spoAdminUrl,
      filter: `HubSiteId eq '${args.options.id}'`,
      fields: 'Title,SiteUrl,SiteId'
    };

    return Cli
      .executeCommandWithOutput(SpoListItemListCommand as Command, { options: { ...options, _: [] } });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let hubSite: HubSite;

    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<HubSite> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/hubsites/getbyid('${encodeURIComponent(args.options.id)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: HubSite): Promise<CommandOutput | void> => {
        hubSite = res;

        if (args.options.includeAssociatedSites !== true || args.options.output && args.options.output !== 'json') {
          return Promise.resolve();
        }

        return spo
          .getSpoAdminUrl(logger, this.debug)
          .then((spoAdminUrl: string): Promise<CommandOutput> => {
            return this.getAssociatedSites(spoAdminUrl, logger, args);
          });
      })
      .then((associatedSitesCommandOutput: CommandOutput | void): void => {
        if (associatedSitesCommandOutput) {
          const associatedSites = JSON.parse((associatedSitesCommandOutput as CommandOutput).stdout) as AssociatedSite[];
          hubSite.AssociatedSites = associatedSites.filter(s => s.SiteId !== args.options.id);
        }

        logger.log(hubSite);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '--includeAssociatedSites'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoHubSiteGetCommand();