import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';

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
  role: string;
}

class EntraM365GroupUserSetCommand extends GraphCommand {
  private readonly allowedRoles: string[] = ['owner', 'member'];

  public get name(): string {
    return commands.M365GROUP_USER_SET;
  }

  public get description(): string {
    return 'Updates role of the specified user in the specified Microsoft 365 Group or Microsoft Teams team';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
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
        option: '-r, --role <role>',
        autocomplete: this.allowedRoles
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `'${args.options.teamId}' is not a valid GUID for option 'teamId'.`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
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

        if (args.options.role && !this.allowedRoles.some(role => role.toLowerCase() === args.options.role.toLowerCase())) {
          return `'${args.options.role}' is not a valid role. Allowed values are: ${this.allowedRoles.join(',')}`;
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
    if (args.options.userName) {
      await this.warn(logger, `Option 'userName' is deprecated. Please use 'ids' or 'userNames' instead.`);
    }

    try {
      const userNames = args.options.userNames || args.options.userName;
      const groupId: string = await this.getGroupId(logger, args);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
      }

      const userIds: string[] = await this.getUserIds(logger, args.options.ids, userNames);

      // we can't simply switch the role
      // first add users to the new role
      await this.addUsers(groupId, userIds, args.options.role);

      // remove users from the old role
      await this.removeUsersFromRole(logger, groupId, userIds, args.options.role);
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
      return formatting.splitAndTrim(userIds);
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving user ID(s) by username(s)...');
    }

    return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames!));
  }

  private async removeUsersFromRole(logger: Logger, groupId: string, userIds: string[], role: string): Promise<void> {
    const userIdsToRemove: string[] = [];
    const currentRole = (role.toLowerCase() === 'member') ? 'owners' : 'members';

    if (this.verbose) {
      await logger.logToStderr(`Removing users from the old role '${currentRole}'.`);
    }

    for (let i = 0; i < userIds.length; i += 20) {
      const userIdsBatch = userIds.slice(i, i + 20);
      const requestOptions = this.getRequestOptions();

      userIdsBatch.map(userId => {
        requestOptions.data.requests.push({
          id: userId,
          method: 'GET',
          url: `/groups/${groupId}/${currentRole}/$count?$filter=id eq '${userId}'`,
          headers: {
            'ConsistencyLevel': 'eventual'
          }
        });
      });

      // send batch request
      const res = await request.post<{ responses: { id: string, status: number; body: any }[] }>(requestOptions);
      for (const response of res.responses) {
        if (response.status === 200) {
          if (response.body === 1) {
            // user can be removed from current role
            userIdsToRemove.push(response.id);
          }
        }
        else {
          throw response.body;
        }
      }
    }

    for (let i = 0; i < userIdsToRemove.length; i += 20) {
      const userIdsBatch = userIds.slice(i, i + 20);
      const requestOptions = this.getRequestOptions();

      userIdsBatch.map(userId => {
        requestOptions.data.requests.push({
          id: userId,
          method: 'DELETE',
          url: `/groups/${groupId}/${currentRole}/${userId}/$ref`
        });
      });

      const res = await request.post<{ responses: { id: string, status: number; body: any }[] }>(requestOptions);
      for (const response of res.responses) {
        if (response.status !== 204) {
          throw response.body;
        }
      }
    }
  }

  private async addUsers(groupId: string, userIds: string[], role: string): Promise<void> {
    for (let i = 0; i < userIds.length; i += 400) {
      const userIdsBatch = userIds.slice(i, i + 400);
      const requestOptions = this.getRequestOptions();

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
            [`${role.toLowerCase() === 'member' ? 'members' : 'owners'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
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

  private getRequestOptions(): CliRequestOptions {
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

    return requestOptions;
  }
}

export default new EntraM365GroupUserSetCommand();