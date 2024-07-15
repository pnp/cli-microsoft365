import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { User } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  newDisplayName?: string;
  description?: string;
  mailNickname?: string;
  ownerIds?: string;
  ownerUserNames?: string;
  memberIds?: string;
  memberUserNames?: string;
  visibility?: string;
}

class EntraGroupSetCommand extends GraphCommand {
  private readonly allowedVisibility: string[] = ['Public', 'Private'];

  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Entra group';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor(){
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
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        newDisplayName: typeof args.options.newDisplayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
        ownerIds: typeof args.options.ownerIds !== 'undefined',
        ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
        memberIds: typeof args.options.memberIds !== 'undefined',
        memberUserNames: typeof args.options.memberUserNames !== 'undefined',
        visibility: typeof args.options.visibility !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '--mailNickname [mailNickname]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '--newDisplayName [newDisplayName]'
      },
      {
        option: '--description [description]'
      },     
      {
        option: '--ownerIds [ownerIds]'
      },
      {
        option: '--ownerUserNames [ownerUserNames]'
      },
      {
        option: '--memberIds [memberIds]'
      },
      {
        option: '--memberUserNames [memberUserNames]'
      },
      {
        option: '--visibility [visibility]',
        autocomplete: this.allowedVisibility
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `Value '${args.options.id}' is not a valid GUID for 'id' option`;
        }

        if (args.options.newDisplayName && args.options.newDisplayName.length > 256) {
          return `The maximum amount of characters for 'newDisplayName' is 256.`;
        }

        if (args.options.mailNickname) {
          if (!validation.isValidMailNickname(args.options.mailNickname)) {
            return `Value '${args.options.mailNickname}' for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \ [] " ; : <> , SPACE.`;
          }

          if (args.options.mailNickname.length > 64) {
            return `The maximum amount of characters for 'mailNickname' is 64.`;
          }
        }

        if (args.options.ownerIds) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ownerIds);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'ownerIds': ${isValidGUIDArrayResult}.`;
          }
        }

        if (args.options.ownerUserNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'ownerUserNames': ${isValidUPNArrayResult}.`;
          }
        }

        if (args.options.memberIds) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.memberIds);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'memberIds': ${isValidGUIDArrayResult}.`;
          }
        }

        if (args.options.memberUserNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.memberUserNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'memberUserNames': ${isValidUPNArrayResult}.`;
          }
        }

        if (args.options.visibility && this.allowedVisibility.indexOf(args.options.visibility) === -1) {
          return `Option 'visibility' must be one of the following values: Public, Private.`;
        }

        if (args.options.newDisplayName === undefined && args.options.description === undefined && args.options.visibility === undefined
          && args.options.ownerIds === undefined && args.options.ownerUserNames === undefined && args.options.memberIds === undefined
          && args.options.memberUserNames === undefined) {
          return `Specify at least one of the following options: 'newDisplayName', 'description', 'visibility', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames'`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'displayName'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('displayName', 'newDisplayName', 'description', 'mailNickname', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'visibility');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let groupId = args.options.id;
    let ownerIds: string[] = [];
    let memberIds: string[] = [];

    try {
      if (args.options.displayName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving group id...`);
        }

        groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
      }

      const requestBody = this.createRequestBody(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${groupId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: requestBody
      };

      await request.patch(requestOptions);

      ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      if (ownerIds.length !== 0) {
        await this.updateUsers(groupId!, 'owners', ownerIds);
      }
      else if (this.verbose) {
        await logger.logToStderr(`No owners to update.`);
      }

      memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);
      if (memberIds.length !== 0) {
        await this.updateUsers(groupId!, 'members', memberIds);
      }
      else if (this.verbose) {
        await logger.logToStderr(`No members to update.`);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  };

  private createRequestBody(options: Options): any {
    const requestBody: any = {
      displayName: options.newDisplayName,
      description: options.description,
      mailNickName: options.mailNickname,
      visibility: options.visibility
    };

    this.addUnknownOptionsToPayload(requestBody, options);
    return requestBody;
  }

  private async getUserIds(logger: Logger, userIds?: string, userNames?: string): Promise<string[]> {
    if (userIds) {
      return formatting.splitAndTrim(userIds);
    }

    if (userNames) {
      return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
    }

    return [];
  }

  private async updateUsers(groupId: string, role: string, userIds: string[]): Promise<void> {
    const groupUsers = await odata.getAllItems<User>(`${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=id`);
    const userIdsToAdd = userIds.filter(userId => groupUsers.findIndex(groupUser => groupUser.id === userId) === -1);
    const userIdsToRemove = groupUsers.filter(groupUser => userIds.findIndex(userId => groupUser.id === userId) === -1);

    for (let i = 0; i < userIdsToAdd.length; i += 400) {
      const userIdsBatch = userIdsToAdd.slice(i, i + 400);
      const requestOptions = this.getRequestOptions();

      // only 20 requests per one batch are allowed
      for (let j = 0; j < userIdsBatch.length; j += 20) {
        // only 20 users can be added in one request
        const userIdsChunk = userIdsBatch.slice(j, j + 20);
        requestOptions.data.requests.push({
          id: j + 1,
          method: 'PATCH',
          url: `/groups/${groupId}`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          body: {
            [`${role}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
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

    for (let i = 0; i < userIdsToRemove.length; i += 20) {
      const userIdsBatch = userIdsToRemove.slice(i, i + 20);
      const requestOptions = this.getRequestOptions();

      userIdsBatch.map(userId => {
        requestOptions.data.requests.push({
          id: userId,
          method: 'DELETE',
          url: `/groups/${groupId}/${role}/${userId}/$ref`
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

export default new EntraGroupSetCommand();