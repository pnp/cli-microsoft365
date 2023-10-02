import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
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
}

class AadGroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_USER_LIST;
  }

  public get description(): string {
    return 'Lists users of a specific Azure AD group';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName', 'userType'];
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
        role: typeof args.options.role !== 'undefined'
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
      const groupId = await this.getGroupId(args.options);
      let users = await this.getOwners(groupId, logger);

      if (args.options.role !== 'Owner') {
        const membersAndGuests = await this.getMembersAndGuests(groupId, logger);
        users = users.concat(membersAndGuests);
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

  private async getGroupId(options: Options): Promise<string> {
    const { groupId, groupDisplayName } = options;

    if (groupId) {
      return groupId;
    }

    return await aadGroup.getGroupIdByDisplayName(groupDisplayName!);
  }

  private async getOwners(groupId: string, logger: Logger): Promise<User[]> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving owners of the group with id ${groupId}`);
    }

    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    const users = await odata.getAllItems<User>(endpoint);

    // Currently there is a bug in the Microsoft Graph that returns Owners as
    // userType 'member'. We therefore update all returned user as owner
    users.forEach(user => {
      user.userType = 'Owner';
    });

    return users;
  }

  private async getMembersAndGuests(groupId: string, logger: Logger): Promise<User[]> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving members of the group with id ${groupId}`);
    }

    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return await odata.getAllItems<User>(endpoint);
  }
}

export default new AadGroupUserListCommand();