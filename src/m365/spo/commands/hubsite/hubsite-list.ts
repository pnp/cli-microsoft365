import request from '../../../../request';
import commands from '../../commands';
import { CommandOption } from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { HubSite } from './HubSite';
import { QueryListResult } from './QueryListResult';
import { AssociatedSite } from './AssociatedSite';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  includeAssociatedSites?: boolean;
}

class SpoHubSiteListCommand extends SpoCommand {
  private batchSize: number = 30;

  public get name(): string {
    return `${commands.HUBSITE_LIST}`;
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let hubSites: HubSite[];
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<{ value: HubSite[]; }> => {
        spoAdminUrl = _spoAdminUrl;

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/hubsites`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
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
            cmd.log('Retrieving associated sites...');
            cmd.log('');
          }
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true,
          body: {
            parameters: {
              ViewXml: "<View><Query><Where><And><And><IsNull><FieldRef Name=\"TimeDeleted\"/></IsNull><Neq><FieldRef Name=\"State\"/><Value Type='Integer'>0</Value></Neq></And><Neq><FieldRef Name=\"HubSiteId\"/><Value Type='Text'>{00000000-0000-0000-0000-000000000000}</Value></Neq></And></Where><OrderBy><FieldRef Name='Title' Ascending='true' /></OrderBy></Query><ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"SiteUrl\"/><FieldRef Name=\"SiteId\"/><FieldRef Name=\"HubSiteId\"/></ViewFields><RowLimit Paged=\"TRUE\">" + this.batchSize + "</RowLimit></View>",
              DatesInUtc: true
            }
          }
        };

        if (this.debug) {
          cmd.log(`Will retrieve associated sites (including the hub sites) in batches of ${this.batchSize}`);
        }

        return this.getSites(requestOptions, requestOptions.url, cmd);
      })
      .then((res: AssociatedSite[] | void): void => {
        if (res) {
          hubSites.forEach(h => {
            const filteredSites = res.filter(f => {
              // Only include sites of which the Site Id is not the same as the
              // Hub Site ID (as this site is the actual hub site) and of which the
              // Hub Site ID matches the ID of the Hub
              return f.SiteId !== f.HubSiteId
                && (f.HubSiteId as string).toUpperCase() == `{${h.ID.toUpperCase()}}`;
            });
            h.AssociatedSites = filteredSites.map(a => {
              return {
                Title: a.Title,
                SiteUrl: a.SiteUrl
              }
            })
          });
        };

        if (args.options.output === 'json') {
          cmd.log(hubSites);
        }
        else {
          cmd.log(hubSites.map(h => {
            return {
              ID: h.ID,
              SiteUrl: h.SiteUrl,
              Title: h.Title
            };
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getSites(reqOptions: any, nonPagedUrl: string, cmd: CommandInstance, sites: AssociatedSite[] = [], batchNumber: number = 0): Promise<AssociatedSite[]> {
    return new Promise<AssociatedSite[]>((resolve: (associatedSites: AssociatedSite[]) => void, reject: (error: any) => void): void => {
      request
        .post<QueryListResult>(reqOptions)
        .then((res: QueryListResult): void => {
          batchNumber++;
          const retrievedSites: AssociatedSite[] = res.Row.length > 0 ? sites.concat(res.Row) : sites;

          if (this.debug) {
            cmd.log(res);
            cmd.log(`Retrieved ${res.Row.length} sites in batch ${batchNumber}`);
          }

          if (!!res.NextHref) {
            reqOptions.url = nonPagedUrl + res.NextHref;
            if (this.debug) {
              cmd.log(`Url for next batch of sites: ${reqOptions.url}`);
            }

            this
              .getSites(reqOptions, nonPagedUrl, cmd, retrievedSites, batchNumber)
              .then((associatedSites: AssociatedSite[]): void => {
                resolve(associatedSites);
              }, (err: any): void => {
                reject(err);
              });
          }
          else {
            if (this.debug) {
              cmd.log(`Retrieved ${retrievedSites.length} sites in total`);
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
        option: '-i, --includeAssociatedSites',
        description: `Include the associated sites in the result (only in JSON output)`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on a SharePoint API that is currently
    in preview and is subject to change once the API reached general
    availability.

    When using the text output type (default), the command lists only the
    values of the ${chalk.grey('ID')}, ${chalk.grey('SiteUrl')} and ${chalk.grey('Title')} properties of the hub site. When setting
    the output type to JSON, all available properties are included in
    the command output.

  Examples:
  
    List hub sites in the current tenant
      ${this.name}

    List hub sites, including their associated sites, in the current tenant. Associated site info is only shown in JSON output.
      ${this.name} --includeAssociatedSites --output json

  More information:

    SharePoint hub sites new in Microsoft 365
      https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547
`);
  }
}

module.exports = new SpoHubSiteListCommand();