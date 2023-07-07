import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { AssociatedSite } from './AssociatedSite';
import { HubSite } from './HubSite';
import { QueryListResult } from './QueryListResult';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  includeAssociatedSites?: boolean;
}

class SpoHubSiteListCommand extends SpoCommand {
  private batchSize: number = 30;

  public get name(): string {
    return commands.HUBSITE_LIST;
  }

  public get description(): string {
    return 'Lists hub sites in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteUrl', 'Title'];
  }

  constructor() {
    super();
    this.batchSize = 30;

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        includeAssociatedSites: args.options.includeAssociatedSites === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --includeAssociatedSites'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      const hubSitesResult = await odata.getAllItems<HubSite>(`${spoAdminUrl}/_api/hubsites`);
      const hubSites = hubSitesResult;

      if (!(args.options.includeAssociatedSites !== true || args.options.output && args.options.output !== 'json')) {
        if (this.debug) {
          logger.logToStderr('Retrieving associated sites...');
          logger.logToStderr('');
        }

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            parameters: {
              ViewXml: "<View><Query><Where><And><And><IsNull><FieldRef Name=\"TimeDeleted\"/></IsNull><Neq><FieldRef Name=\"State\"/><Value Type='Integer'>0</Value></Neq></And><Neq><FieldRef Name=\"HubSiteId\"/><Value Type='Text'>{00000000-0000-0000-0000-000000000000}</Value></Neq></And></Where><OrderBy><FieldRef Name='Title' Ascending='true' /></OrderBy></Query><ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"SiteUrl\"/><FieldRef Name=\"SiteId\"/><FieldRef Name=\"HubSiteId\"/></ViewFields><RowLimit Paged=\"TRUE\">" + this.batchSize + "</RowLimit></View>",
              DatesInUtc: true
            }
          }
        };

        if (this.debug) {
          logger.logToStderr(`Will retrieve associated sites (including the hub sites) in batches of ${this.batchSize}`);
        }

        const res = await this.getSites(requestOptions, requestOptions.url as string, logger);

        if (res) {
          hubSites.forEach(h => {
            const filteredSites = res.filter(f => {
              // Only include sites of which the Site Id is not the same as the
              // Hub Site ID (as this site is the actual hub site) and of which the
              // Hub Site ID matches the ID of the Hub
              return f.SiteId !== f.HubSiteId
                && (f.HubSiteId as string).toUpperCase() === `{${h.ID.toUpperCase()}}`;
            });
            h.AssociatedSites = filteredSites.map(a => {
              return {
                Title: a.Title,
                SiteUrl: a.SiteUrl
              };
            });
          });
        }
      }

      logger.log(hubSites);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSites(reqOptions: any, nonPagedUrl: string, logger: Logger, sites: AssociatedSite[] = [], batchNumber: number = 0): Promise<AssociatedSite[]> {
    const res = await request.post<QueryListResult>(reqOptions);
    batchNumber++;
    const retrievedSites: AssociatedSite[] = res.Row.length > 0 ? sites.concat(res.Row) : sites;

    if (this.debug) {
      logger.logToStderr(res);
      logger.logToStderr(`Retrieved ${res.Row.length} sites in batch ${batchNumber}`);
    }

    if (!!res.NextHref) {
      reqOptions.url = nonPagedUrl + res.NextHref;
      if (this.debug) {
        logger.logToStderr(`Url for next batch of sites: ${reqOptions.url}`);
      }
      return this.getSites(reqOptions, nonPagedUrl, logger, retrievedSites, batchNumber);
    }
    else {
      if (this.debug) {
        logger.logToStderr(`Retrieved ${retrievedSites.length} sites in total`);
      }

      return retrievedSites;
    }
  }
}

module.exports = new SpoHubSiteListCommand();