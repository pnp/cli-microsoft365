import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils';
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

  constructor() {
    super();
    this.batchSize = 30;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.includeAssociatedSites = args.options.includeAssociatedSites === true;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteUrl', 'Title'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let hubSites: HubSite[];
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<{ value: HubSite[]; }> => {
        spoAdminUrl = _spoAdminUrl;

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/hubsites`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: { value: HubSite[] }): Promise<any[] | void> => {
        hubSites = res.value;

        if (args.options.includeAssociatedSites !== true || args.options.output !== 'json') {
          return Promise.resolve();
        }
        else {
          if (this.debug) {
            logger.logToStderr('Retrieving associated sites...');
            logger.logToStderr('');
          }
        }

        const requestOptions: any = {
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

        return this.getSites(requestOptions, requestOptions.url, logger);
      })
      .then((res: AssociatedSite[] | void): void => {
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

        logger.log(hubSites);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getSites(reqOptions: any, nonPagedUrl: string, logger: Logger, sites: AssociatedSite[] = [], batchNumber: number = 0): Promise<AssociatedSite[]> {
    return new Promise<AssociatedSite[]>((resolve: (associatedSites: AssociatedSite[]) => void, reject: (error: any) => void): void => {
      request
        .post<QueryListResult>(reqOptions)
        .then((res: QueryListResult): void => {
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

            this
              .getSites(reqOptions, nonPagedUrl, logger, retrievedSites, batchNumber)
              .then((associatedSites: AssociatedSite[]): void => {
                resolve(associatedSites);
              }, (err: any): void => {
                reject(err);
              });
          }
          else {
            if (this.debug) {
              logger.logToStderr(`Retrieved ${retrievedSites.length} sites in total`);
            }

            resolve(retrievedSites);
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --includeAssociatedSites'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoHubSiteListCommand();