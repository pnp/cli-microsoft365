import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import teamsCommands from '../../../teams/commands.js';
import aadCommands from '../../aadCommands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName?: string;
  ids?: string;
  userNames?: string;
  groupId?: string;
  groupName?: string;
  teamId?: string;
  teamName?: string;
  role?: string;
}

class EntraM365GroupUserAddCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Adds user to specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.USER_ADD, aadCommands.M365GROUP_USER_ADD];
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
        role: args.options.role !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        ids: typeof args.options.ids !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '-r, --role [role]',
        autocomplete: ['Owner', 'Member']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.ids && args.options.ids.split(',').some(e => !validation.isValidGuid(e))) {
          return `${args.options.ids} contains one or more invalid GUIDs`;
        }

        if (args.options.userNames && args.options.userNames.split(',').some(e => !validation.isValidUserPrincipalName(e))) {
          return `${args.options.userNames} contains one or more invalid usernames`;
        }

        if (args.options.role) {
          if (['owner', 'member'].indexOf(args.options.role.toLowerCase()) === -1) {
            return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupId', 'groupName', 'teamId', 'teamName'] });
    this.optionSets.push({ options: ['userName', 'ids', 'userNames'] });
  }

  #initTypes(): void {
    this.types.string.push('userName', 'ids', 'userNames', 'groupId', 'groupName', 'teamId', 'teamName', 'role');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_USER_ADD, commands.M365GROUP_USER_ADD);

    if (args.options.userName) {
      args.options.userNames = args.options.userName;

      this.warn(logger, `Option 'userName' is deprecated. Please use 'ids' or 'userNames' instead.`);
    }

    try {
      const providedGroupId: string = await this.getGroupId(logger, args);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(providedGroupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${providedGroupId}' is not a Microsoft 365 group.`);
      }

      const userIds: string[] = await this.getUserIds(logger, args.options.ids, args.options.userNames);

      if (this.verbose) {
        await logger.logToStderr(`Adding user(s) ${args.options.ids || args.options.userNames} to group ${args.options.groupId || args.options.groupName || args.options.teamId || args.options.teamName}...`);
      }

      await this.addUsers(providedGroupId, userIds, args.options.role);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.groupId) {
      return args.options.groupId;
    }

    if (args.options.teamId) {
      return args.options.teamId;
    }

    const name = args.options.groupName || args.options.teamName;

    if (this.verbose) {
      await logger.logToStderr('Retrieving Group ID by display name...');
    }

    return entraGroup.getGroupIdByDisplayName(name!);
  }

  private async getUserIds(logger: Logger, userIds: string | undefined, userNames: string | undefined): Promise<string[]> {
    if (userIds) {
      return userIds.split(',').map(o => o.trim());
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving user ID(s) by username(s)...');
    }

    return entraUser.getUserIdsByUpns(userNames!.split(',').map(u => u.trim()));
  }

  private async addUsers(groupId: string, userIds: string[], role: string | undefined): Promise<void> {
    for (let i = 0; i < userIds.length; i += 400) {
      const userIdsBatch = userIds.slice(i, i + 400);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/$batch`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          requests: []
        }
      };

      for (let j = 0; j < userIdsBatch.length; j += 20) {
        const userIdsChunk = userIdsBatch.slice(j, j + 20);
        requestOptions.data.requests.push({
          id: j + 1,
          method: 'PATCH',
          url: `/groups/${groupId}`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            [`${((typeof role !== 'undefined') ? role : '').toLowerCase() === 'owner' ? 'owners' : 'members'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
          }
        });
      }

      const res = await request.post<{ responses: { status: number; body: any }[] }>(requestOptions);
      for (const response of res.responses) {
        if (response.status !== 204) {
          throw response.body;
        }
      }
    }
  }
}

export default new EntraM365GroupUserAddCommand();