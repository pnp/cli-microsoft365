import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  role?: string;
  properties?: string;
  filter?: string;
}

interface ExtendedUser extends User {
  roles: string[];
}

class AadGroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_USER_LIST;
  }

  public get description(): string {
    return 'Lists users of a specific Azure AD group';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName', 'roles'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        groupId: typeof args.options.groupId !== 'undefined',
        groupDisplayName: typeof args.options.groupDisplayName !== 'undefined',
        role: typeof args.options.role !== 'undefined',
        properties: typeof args.options.properties !== 'undefined',
        filter: typeof args.options.filter !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "-i, --groupId [groupId]"
      },
      {
        option: "-n, --groupDisplayName [groupDisplayName]"
      },
      {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member"]
      },
      {
        option: "-p, --properties [properties]"
      },
      {
        option: "-f, --filter [filter]"
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['groupId', 'groupDisplayName']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.role) {
          if (['Owner', 'Member'].indexOf(args.options.role) === -1) {
            return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args.options);

      let users: User[] = [];

      switch (args.options.role) {
        case 'Owner':
          users = await this.getOwners(args.options, groupId, logger);
          break;
        case 'Member':
          users = await this.getMembers(args.options, groupId, logger);
          break;
        default:
          const owners = await this.getOwners(args.options, groupId, logger);
          const members = await this.getMembers(args.options, groupId, logger);

          if (!args.options.properties) {
            owners.forEach((owner: ExtendedUser) => {
              for (let i = 0; i < members.length; i++) {
                if (members[i].userPrincipalName === owner.userPrincipalName) {
                  if (!owner.roles.includes('Member')) {
                    owner.roles.push('Member');
                  }
                }
              }
            });
          }

          users = owners.concat(members);

          if (!args.options.properties) {
            users = users.filter((value, index, array) => index === array.findIndex(item => item.userPrincipalName === value.userPrincipalName));
          }
      }

      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    const { groupId, groupDisplayName } = options;

    if (groupId) {
      return groupId;
    }

    return await aadGroup.getGroupIdByDisplayName(groupDisplayName!);
  }

  private async getOwners(options: Options, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving owners of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `?$select=${properties.split(',').map(p => formatting.encodeQueryParameter(p.trim())).join(',')}` :
      '?$select=id,displayName,userPrincipalName,givenName,surname';
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners${selectProperties}`;

    let owners: ExtendedUser[] = [];

    if (filter) {
      // While using the filter, we need to specify the ConsistencyLevel header.
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}&$filter=${encodeURIComponent(filter)}&$count=true`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          ConsistencyLevel: 'eventual'
        },
        responseType: 'json'
      };

      owners = await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      owners = await odata.getAllItems<ExtendedUser>(endpoint);
    }

    if (!properties) {
      owners.forEach((user: ExtendedUser) => {
        user.roles = ['Owner'];
      });
    }

    return owners;
  }

  private async getMembers(options: Options, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving members of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `?$select=${properties.split(',').map(p => formatting.encodeQueryParameter(p.trim())).join(',')}` :
      '?$select=id,displayName,userPrincipalName,givenName,surname';
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members${selectProperties}`;

    let members: ExtendedUser[] = [];

    if (filter) {
      // While using the filter, we need to specify the ConsistencyLevel header.
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}&$filter=${encodeURIComponent(filter)}&$count=true`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          ConsistencyLevel: 'eventual'
        },
        responseType: 'json'
      };

      members = await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      members = await odata.getAllItems<ExtendedUser>(endpoint);
    }

    if (!properties) {
      members.forEach((user: ExtendedUser) => {
        user.roles = ['Member'];
      });
    }

    return members;
  }
}

export default new AadGroupUserListCommand();