import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { HubSite } from './HubSite';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
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
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
      .then((res: any): void => {
        logger.log(res);
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
      { option: '-u, --url [url]' }
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