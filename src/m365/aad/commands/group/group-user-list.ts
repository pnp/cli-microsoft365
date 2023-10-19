import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
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

      const users: ExtendedUser[] = [];

      if (!args.options.role || args.options.role === 'Owner') {
        const owners = await this.getUsers(args.options, 'Owners', groupId, logger);
        owners.forEach(owner => users.push({ ...owner, roles: ['Owner'] }));
      }

      if (!args.options.role || args.options.role === 'Member') {
        const members = await this.getUsers(args.options, 'Members', groupId, logger);

        members.forEach((member: ExtendedUser) => {
          const user = users.find((u: ExtendedUser) => u.id === member.id);

          if (user !== undefined) {
            user.roles.push('Member');
          }
          else {
            users.push({ ...member, roles: ['Member'] });
          }
        });
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
      await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
      'id,displayName,userPrincipalName,givenName,surname';
    const allSelectProperties: string[] = selectProperties.split(',');
    const propertiesWithSlash: string[] = allSelectProperties.filter(item => item.includes('/'));

    let fieldExpand: string = '';
    propertiesWithSlash.forEach(p => {
      if (fieldExpand.length > 0) {
        fieldExpand += ',';
      }

      fieldExpand += `${p.split('/')[0]}($select=${p.split('/')[1]})`;
    });

    const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';

    const selectParam = allSelectProperties.filter(item => !item.includes('/'));
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=${selectParam}${expandParam}`;

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

    return users;
  }
}

export default new AadGroupUserListCommand();