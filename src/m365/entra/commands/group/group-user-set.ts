import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  ids?: string;
  userNames?: string;
  role: string;
}

class EntraGroupUserSetCommand extends GraphCommand {
  private readonly roleValues = ['Owner', 'Member'];

  public get name(): string {
    return commands.GROUP_USER_SET;
  }

  public get description(): string {
    return 'Updates role of users in a Microsoft Entra ID group';
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupDisplayName: typeof args.options.groupDisplayName !== 'undefined',
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
        option: '-n, --groupDisplayName [groupDisplayName]'
      },
      {
        option: '--ids [ids]'
      },
      {
        option: '--userNames [userNames]'
      },
      {
        option: '-r, --role <role>',
        autocomplete: this.roleValues
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID for option groupId.`;
        }

        if (args.options.ids) {
          const ids = args.options.ids.split(',').map(i => i.trim());
          if (!validation.isValidGuidArray(ids)) {
            const invalidGuid = ids.find(id => !validation.isValidGuid(id));
            return `'${invalidGuid}' is not a valid GUID for option 'ids'.`;
          }
        }

        if (args.options.userNames) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.userNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `User principal name '${isValidUserPrincipalNameArray}' is invalid for option 'userNames'.`;
          }
        }

        if (this.roleValues.indexOf(args.options.role) === -1) {
          return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupDisplayName'] },
      { options: ['ids', 'userNames'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('groupId', 'groupDisplayName', 'ids', 'userNames', 'role');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding user(s) ${args.options.ids || args.options.userNames} to group ${args.options.groupId || args.options.groupDisplayName}...`);
      }

      const groupId = await this.getGroupId(logger, args.options);
      const userIds = await this.getUserIds(logger, args.options);

      // we can't simply switch the role
      // first add users to the new role
      await this.addUsers(groupId, userIds, args.options);

      // remove users from the old role
      await this.removeUsersFromRole(logger, groupId, userIds, args.options);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(logger: Logger, options: Options): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ID of group ${options.groupDisplayName}...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupDisplayName!);
  }

  private async getUserIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.ids) {
      return options.ids.split(',').map(i => i.trim());
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of user(s)...');
    }

    return entraUser.getUserIdsByUpns(options.userNames!.split(',').map(u => u.trim()));
  }

  private async removeUsersFromRole(logger: Logger, groupId: string, userIds: string[], options: Options): Promise<void> {
    const userIdsToRemove: string[] = [];
    const currentRole = options.role === 'Member' ? 'owners' : 'members';

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

  private async addUsers(groupId: string, userIds: string[], options: Options): Promise<void> {
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
            [`${options.role === 'Member' ? 'members' : 'owners'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
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

export default new EntraGroupUserSetCommand();