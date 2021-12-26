import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { GroupExtended } from './GroupExtended';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName?: string;
  mailNickname?: string;
  includeSiteUrl: boolean;
  deleted?: boolean;
  orphaned?: boolean;
}

class AadO365GroupListCommand extends GraphItemsListCommand<GroupExtended> {
  public get name(): string {
    return commands.O365GROUP_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft 365 Groups in the current tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    telemetryProps.mailNickname = typeof args.options.mailNickname !== 'undefined';
    telemetryProps.includeSiteUrl = args.options.includeSiteUrl;
    telemetryProps.deleted = args.options.deleted;
    telemetryProps.orphaned = args.options.orphaned;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mailNickname', 'deletedDateTime', 'siteUrl'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const groupFilter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const displayNameFilter: string = args.options.displayName ? ` and startswith(DisplayName,'${encodeURIComponent(args.options.displayName).replace(/'/g, `''`)}')` : '';
    const mailNicknameFilter: string = args.options.mailNickname ? ` and startswith(MailNickname,'${encodeURIComponent(args.options.mailNickname).replace(/'/g, `''`)}')` : '';
    const expandOwners: string = args.options.orphaned ? '&$expand=owners' : '';
    const topCount: string = '&$top=100';

    let endpoint: string = `${this.resource}/v1.0/groups${groupFilter}${displayNameFilter}${mailNicknameFilter}${expandOwners}${topCount}`;

    if (args.options.deleted) {
      endpoint = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${groupFilter}${displayNameFilter}${mailNicknameFilter}${topCount}`;
    }

    this
      .getAllItems(endpoint, logger, true)
      .then((): Promise<any> => {
        if (args.options.orphaned) {
          const orphanedGroups: GroupExtended[] = [];

          this.items.forEach((group) => {
            if (!group.owners || group.owners.length === 0) {
              orphanedGroups.push(group);
            }
          });

          this.items = orphanedGroups;
        }

        if (args.options.includeSiteUrl) {
          return Promise.all(this.items.map(g => this.getGroupSiteUrl(g.id as string)));
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

        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getGroupSiteUrl(groupId: string): Promise<{ id: string, url: string }> {
    return new Promise<{ id: string, url: string }>((resolve: (siteInfo: { id: string, url: string }) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/groups/${groupId}/drive?$select=webUrl`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .get<{ webUrl: string }>(requestOptions)
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
        option: '-d, --displayName [displayName]'
      },
      {
        option: '-m, --mailNickname [displayName]'
      },
      {
        option: '--includeSiteUrl'
      },
      {
        option: '--deleted'
      },
      {
        option: '--orphaned'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.deleted && args.options.includeSiteUrl) {
      return 'You can\'t retrieve site URLs of deleted Microsoft 365 Groups';
    }

    return true;
  }
}

module.exports = new AadO365GroupListCommand();