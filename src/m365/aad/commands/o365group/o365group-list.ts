import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const groupFilter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const displayNameFilter: string = args.options.displayName ? ` and startswith(DisplayName,'${formatting.encodeQueryParameter(args.options.displayName)}')` : '';
    const mailNicknameFilter: string = args.options.mailNickname ? ` and startswith(MailNickname,'${formatting.encodeQueryParameter(args.options.mailNickname)}')` : '';
    const expandOwners: string = args.options.orphaned ? '&$expand=owners' : '';
    const topCount: string = '&$top=100';

    let endpoint: string = `${this.resource}/v1.0/groups${groupFilter}${displayNameFilter}${mailNicknameFilter}${expandOwners}${topCount}`;

    if (args.options.deleted) {
      endpoint = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${groupFilter}${displayNameFilter}${mailNicknameFilter}${topCount}`;
    }

    try {
      let groups: GroupExtended[] = [];
      groups = await odata.getAllItems<GroupExtended>(endpoint);

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
        const res = await Promise.all(groups.map(g => this.getGroupSiteUrl(g.id as string)));
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
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupSiteUrl(groupId: string): Promise<{ id: string, url: string }> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/groups/${groupId}/drive?$select=webUrl`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ webUrl: string }>(requestOptions);
    return {
      id: groupId,
      url: res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : ''
    };
  }
}

module.exports = new AadO365GroupListCommand();