import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  groupId?: string;
  groupName?: string;
  userName?: string;
  ids?: string;
  userNames?: string;
  force?: boolean;
}

class EntraM365GroupUserRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force,
        teamId: typeof args.options.teamId !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        ids: typeof args.options.ids !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--ids [ids]'
      },
      {
        option: '--userNames [userNames]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID for option 'teamId'.`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID for option 'groupId'.`;
        }

        if (args.options.ids) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ids);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
          }
        }

        if (args.options.userNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.userNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'userNames': ${isValidUPNArrayResult}.`;
          }
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `The specified userName '${args.options.userName}' is not a valid user principal name.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['groupId', 'teamId', 'groupName', 'teamName']
      },
      {
        options: ['userName', 'ids', 'userNames']
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('groupId', 'groupName', 'teamId', 'teamName', 'userName', 'ids', 'userNames');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const groupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string;

    const removeUser = async (): Promise<void> => {
      try {
        const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

        if (!isUnifiedGroup) {
          throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
        }

        const userNames = args.options.userNames || args.options.userName;
        const userIds: string[] = await this.getUserIds(logger, args.options.ids, userNames);

        await this.removeUsersFromGroup(groupId, userIds, 'owners');
        await this.removeUsersFromGroup(groupId, userIds, 'members');
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeUser();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${args.options.userName || args.options.userNames || args.options.ids} from ${args.options.groupId || args.options.groupName || args.options.teamId || args.options.teamName}?` });

      if (result) {
        await removeUser();
      }
    }
  }

  private async getUserIds(logger: Logger, userIds?: string, userNames?: string): Promise<string[]> {
    if (userIds) {
      return formatting.splitAndTrim(userIds);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving user IDs for {userNames}...`);
    }

    return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames!));
  }

  private async removeUsersFromGroup(groupId: string, userIds: string[], role: string): Promise<void> {
    for (const userId of userIds) {
      try {
        await request.delete({
          url: `${this.resource}/v1.0/groups/${groupId}/${role}/${userId}/$ref`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        });
      }
      catch (err: any) {
        // the 404 error is accepted
        if (err.response.status !== 404) {
          throw err.response.data;
        }
      }
    }
  }
}

export default new EntraM365GroupUserRemoveCommand();