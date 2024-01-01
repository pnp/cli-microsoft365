import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filter?: string;
  groupId?: string;
  groupDisplayName?: string;
  properties?: string;
  role?: string;
}

interface ExtendedUser extends User {
  roles: string[];
}

class AadM365GroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_LIST;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft 365 group";
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
        autocomplete: ["Owner", "Member", "Guest"]
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
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.role) {
          if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
            return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.role === 'Guest') {
        this.warn(logger, `Value 'Guest' for the option role is deprecated.`);
      }

      const groupId = await this.getGroupId(args.options, logger);
      const isUnifiedGroup = await aadGroup.isUnifiedGroup(groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group '${args.options.groupId || args.options.groupDisplayName}' is not a Microsoft 365 group.`);
      }

      let users: ExtendedUser[] = [];
      if (!args.options.role || args.options.role === 'Owner') {
        const owners = await this.getUsers(args.options, 'Owners', groupId, logger);
        owners.forEach(owner => users.push({ ...owner, roles: ['Owner'], userType: 'Owner' }));
      }

      if (!args.options.role || args.options.role === 'Member' || args.options.role === 'Guest') {
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

      if (args.options.role) {
        users = users.filter(i => i.userType === args.options.role);
      }

      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(options: Options, logger: Logger): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving Group Id...');
    }

    return await aadGroup.getGroupIdByDisplayName(options.groupDisplayName!);
  }

  private async getUsers(options: Options, role: string, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
      'id,displayName,userPrincipalName,givenName,surname,userType';
    const allSelectProperties: string[] = selectProperties.split(',');
    const propertiesWithSlash: string[] = allSelectProperties.filter(item => item.includes('/'));

    const fieldsToExpand: string[] = [];
    propertiesWithSlash.forEach(p => {
      const propertiesSplit: string[] = p.split('/');
      fieldsToExpand.push(`${propertiesSplit[0]}($select=${propertiesSplit[1]})`);
    });

    const fieldExpand: string = fieldsToExpand.join(',');

    const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
    const selectParam = allSelectProperties.filter(item => !item.includes('/'));
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=${selectParam}${expandParam}`;

    if (filter) {
      // While using the filter, we need to specify the ConsistencyLevel header.
      // Can be refactored when the header is no longer necessary.
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}&$filter=${encodeURIComponent(filter)}&$count=true`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          ConsistencyLevel: 'eventual'
        },
        responseType: 'json'
      };

      return await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      return await odata.getAllItems<ExtendedUser>(endpoint);
    }
  }
}

export default new AadM365GroupUserListCommand();