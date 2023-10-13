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
          users = await this.getUsers(args.options, 'Owner', groupId, logger);
          break;
        case 'Member':
          users = await this.getUsers(args.options, 'Member', groupId, logger);
          break;
        default:
          const owners = await this.getUsers(args.options, 'Owner', groupId, logger);
          const members = await this.getUsers(args.options, 'Member', groupId, logger);

          if (!args.options.properties) {
            owners.forEach((owner: ExtendedUser) => {
              for (let i = 0; i < members.length; i++) {
                if (members[i].id === owner.id) {
                  if (!owner.roles.includes('Member')) {
                    owner.roles.push('Member');
                  }
                }
              }
            });
          }

          users = owners.concat(members);
          users = users.filter((value, index, array) => index === array.findIndex(item => item.id === value.id));
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

  private async getUsers(options: Options, role: string, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ${role}s of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `?$select=${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => formatting.encodeQueryParameter(p.trim())).join(',')}` :
      '?$select=id,displayName,userPrincipalName,givenName,surname';
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${role}s${selectProperties}`;

    let users: ExtendedUser[] = [];

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

      users = await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      users = await odata.getAllItems<ExtendedUser>(endpoint);
    }

    users.forEach((user: ExtendedUser) => {
      user.roles = [role];
    });

    return users;
  }
}

export default new AadGroupUserListCommand();