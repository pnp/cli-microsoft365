import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
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

class AadO365GroupListCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft 365 Groups in the current tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
        includeSiteUrl: args.options.includeSiteUrl,
        deleted: args.options.deleted,
        orphaned: args.options.orphaned
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.deleted && args.options.includeSiteUrl) {
          return 'You can\'t retrieve site URLs of deleted Microsoft 365 Groups';
        }
    
        return true;
      }
    );
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

    let groups: GroupExtended[] = [];

    odata
      .getAllItems<GroupExtended>(endpoint)
      .then((_groups): Promise<any> => {
        groups = _groups;

        if (args.options.orphaned) {
          const orphanedGroups: GroupExtended[] = [];

          groups.forEach((group) => {
            if (!group.owners || group.owners.length === 0) {
              orphanedGroups.push(group);
            }
          });

          groups = orphanedGroups;
        }

        if (args.options.includeSiteUrl) {
          return Promise.all(groups.map(g => this.getGroupSiteUrl(g.id as string)));
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res?: { id: string, url: string }[]): void => {
        if (res) {
          res.forEach(r => {
            for (let i: number = 0; i < groups.length; i++) {
              if (groups[i].id !== r.id) {
                continue;
              }

              groups[i].siteUrl = r.url;
              break;
            }
          });
        }

        logger.log(groups);
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
}

module.exports = new AadO365GroupListCommand();