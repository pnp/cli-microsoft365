import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { HubSite } from './HubSite';
import request, { CliRequestOptions } from '../../../../request';
import { spo } from '../../../../utils/spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  url?: string;
  parentId?: string;
  parentTitle?: string;
  parentUrl?: string;
}

class SpoHubSiteConnectCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_CONNECT;
  }

  public get description(): string {
    return 'Connect a hub site to a parent hub site';
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
        parentId: typeof args.options.parentId !== 'undefined',
        parentTitle: typeof args.options.parentTitle !== 'undefined',
        parentUrl: typeof args.options.parentUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-u, --url [url]'
      },
      {
        option: '--parentId [parentId]'
      },
      {
        option: '--parentTitle [parentTitle]'
      },
      {
        option: '--parentUrl [parentUrl]'
      },
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title', 'url'] },
      { options: ['parentId', 'parentTitle', 'parentUrl'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID`;
        }

        if (args.options.parentId && !validation.isValidGuid(args.options.parentId)) {
          return `'${args.options.parentId}' is not a valid GUID`;
        }

        if (args.options.url) {
          const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);

          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (args.options.parentUrl) {
          const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.parentUrl);

          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Connecting hub site '${args.options.id || args.options.title || args.options.url}' to hub site '${args.options.parentId || args.options.parentTitle || args.options.parentUrl}'...`);
    }

    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const hubSites = await this.getHubSites(spoAdminUrl);

      const hubSite = this.getSpecificHubSite(hubSites, args.options.id, args.options.title, args.options.url);
      const parentHubSite = this.getSpecificHubSite(hubSites, args.options.parentId, args.options.parentTitle, args.options.parentUrl);

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

  private getSpecificHubSite(hubSites: HubSite[], id?: string, title?: string, url?: string): HubSite {
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
      throw `Multiple hub sites with name '${title}' found: ${filteredHubSites.map(s => s.ID).join(',')}.`;
    }

    return filteredHubSites[0];
  }
}

module.exports = new SpoHubSiteConnectCommand();