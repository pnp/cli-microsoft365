import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { Group } from './Group';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName?: string;
  mailNickname?: string;
  includeSiteUrl: boolean;
  deleted?: boolean;
}

class GraphO365GroupListCommand extends GraphItemsListCommand<Group> {
  public get name(): string {
    return `${commands.O365GROUP_LIST}`;
  }

  public get description(): string {
    return 'Lists Office 365 Groups in the current tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    telemetryProps.mailNickname = typeof args.options.mailNickname !== 'undefined';
    telemetryProps.includeSiteUrl = args.options.includeSiteUrl;
    telemetryProps.deleted = args.options.deleted;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const groupFilter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const displayNameFilter: string = args.options.displayName ? ` and startswith(DisplayName,'${encodeURIComponent(args.options.displayName).replace(/'/g, `''`)}')` : '';
    const mailNicknameFilter: string = args.options.mailNickname ? ` and startswith(MailNickname,'${encodeURIComponent(args.options.mailNickname).replace(/'/g, `''`)}')` : '';
    const topCount: string = '&$top=100';
    let endpoint: string = `${auth.service.resource}/v1.0/groups${groupFilter}${displayNameFilter}${mailNicknameFilter}${topCount}`;

    if (args.options.deleted) {
      if (this.debug) {
        cmd.log('Switch to BETA endpoint to retrieve deleted items');
      }

      endpoint = `${auth.service.resource}/beta/directory/deletedItems/Microsoft.Graph.Group${groupFilter}${displayNameFilter}${mailNicknameFilter}${topCount}`;
    }

    this
      .getAllItems(endpoint, cmd, true)
      .then((): Promise<any> => {
        if (args.options.includeSiteUrl) {
          return Promise.all(this.items.map(g => this.getGroupSiteUrl(g.id, cmd)));
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res?: { id: string, url: string }[]): void => {
        if (res) {
          res.forEach(r => {
            for (let i: number = 0; i < this.items.length; i++) {
              if (this.items[i].id !== r.id) {
                continue;
              }

              this.items[i].siteUrl = r.url;
              break;
            }
          });
        }

        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(g => {
            const group: any = {
              id: g.id,
              displayName: g.displayName,
              mailNickname: g.mailNickname
            };

            if (args.options.deleted) {
              group.deletedDateTime = g.deletedDateTime;
            }

            if (args.options.includeSiteUrl) {
              group.siteUrl = g.siteUrl;
            }

            return group;
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getGroupSiteUrl(groupId: string, cmd: CommandInstance): Promise<{ id: string, url: string }> {
    return new Promise<{ id: string, url: string }>((resolve: (siteInfo: { id: string, url: string }) => void, reject: (error: any) => void): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/groups/${groupId}/drive?$select=webUrl`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.get(requestOptions);
        })
        .then((res: { webUrl: string }): void => {
          resolve({
            id: groupId,
            url: res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : ''
          });
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-d, --displayName [displayName]',
        description: 'Retrieve only groups with displayName starting with the specified value'
      },
      {
        option: '-m, --mailNickname [displayName]',
        description: 'Retrieve only groups with mailNickname starting with the specified value'
      },
      {
        option: '--includeSiteUrl',
        description: 'Set to retrieve the site URL for each group'
      },
      {
        option: '--deleted',
        description: 'Set to only retrieve deleted groups'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.deleted && args.options.includeSiteUrl) {
        return 'You can\'t retrieve site URLs of deleted Office 365 Groups';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To list available Office 365 Groups, you have to first connect to
    the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

    Using the ${chalk.blue('--includeSiteUrl')} option, you can retrieve the URL
    of the site associated with the particular Office 365 Group. If you however
    retrieve too many groups and will try to get their site URLs, you will most
    likely get an error as the command will get throttled, issuing too many
    requests, too frequently. If you get an error, consider narrowing down
    the result set using the ${chalk.blue('--displayName')} and ${chalk.blue('--mailNickname')} filters.
    
    Retrieving the URL of the site associated with the particular
    Office 365 Group is not possible when retrieving deleted groups.

  Examples:
  
    List all Office 365 Groups in the tenant
      ${chalk.grey(config.delimiter)} ${this.name}

    List Office 365 Groups with display name starting with ${chalk.grey(`Project`)}
      ${chalk.grey(config.delimiter)} ${this.name} --displayName Project

    List Office 365 Groups mail nick name starting with ${chalk.grey(`team`)}
      ${chalk.grey(config.delimiter)} ${this.name} --mailNickname team

    List deleted Office 365 Groups with display name starting with ${chalk.grey(`Project`)}
      ${chalk.grey(config.delimiter)} ${this.name} --displayName Project --deleted

    List deleted Office 365 Groups with mail nick name starting with ${chalk.grey(`team`)}
      ${chalk.grey(config.delimiter)} ${this.name} --mailNickname team --deleted

    List Office 365 Groups with display name starting with ${chalk.grey(`Project`)} including
    the URL of the corresponding SharePoint site
      ${chalk.grey(config.delimiter)} ${this.name} --displayName Project --includeSiteUrl
`);
  }
}

module.exports = new GraphO365GroupListCommand();