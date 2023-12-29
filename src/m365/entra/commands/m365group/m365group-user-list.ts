import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role?: string;
  groupId: string;
}

class AadM365GroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_LIST;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft 365 group";
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_USER_LIST];
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
        role: args.options.role
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "-i, --groupId <groupId>"
      },
      {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member", "Guest"]
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.groupId as string)) {
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
      const isUnifiedGroup = await aadGroup.isUnifiedGroup(args.options.groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.groupId}' is not a Microsoft 365 group.`);
      }

      let users = await this.getOwners(args.options.groupId, logger);

      if (args.options.role !== 'Owner') {
        const membersAndGuests = await this.getMembersAndGuests(args.options.groupId, logger);
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

export default new AadM365GroupUserListCommand();