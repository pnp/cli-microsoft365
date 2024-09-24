import { Group } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { User } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  newDisplayName?: string;
  description?: string;
  ownerIds?: string;
  ownerUserNames?: string;
  memberIds?: string;
  memberUserNames?: string;
  isPrivate?: boolean;
  logoPath?: string;
  allowExternalSenders?: boolean;
  autoSubscribeNewMembers?: boolean;
  hideFromAddressLists?: boolean;
  hideFromOutlookClients?: boolean;
}

class EntraM365GroupSetCommand extends GraphCommand {
  private static numRepeat: number = 15;
  private pollingInterval: number = 500;

  public get name(): string {
    return commands.M365GROUP_SET;
  }

  public get description(): string {
    return 'Updates Microsoft 365 Group properties';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_SET];
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
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        newDisplayName: typeof args.options.newDisplayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        ownerIds: typeof args.options.ownerIds !== 'undefined',
        ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
        memberIds: typeof args.options.memberIds !== 'undefined',
        memberUserNames: typeof args.options.memberUserNames !== 'undefined',
        isPrivate: !!args.options.isPrivate,
        logoPath: typeof args.options.logoPath !== 'undefined',
        allowExternalSenders: !!args.options.allowExternalSenders,
        autoSubscribeNewMembers: !!args.options.autoSubscribeNewMembers,
        hideFromAddressLists: !!args.options.hideFromAddressLists,
        hideFromOutlookClients: !!args.options.hideFromOutlookClients
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '--newDisplayName [newDisplayName]'
      },
      {
        option: '-d, --description [description]'
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
        option: '--isPrivate [isPrivate]',
        autocomplete: ['true', 'false']
      },
      {
        option: '-l, --logoPath [logoPath]'
      },
      {
        option: '--allowExternalSenders [allowExternalSenders]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--autoSubscribeNewMembers [autoSubscribeNewMembers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hideFromAddressLists [hideFromAddressLists]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hideFromOutlookClients [hideFromOutlookClients]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName'] });
    this.optionSets.push({
      options: ['ownerIds', 'ownerUserNames'],
      runsWhen: (args) => {
        return args.options.ownerIds !== undefined || args.options.ownerUserNames !== undefined;
      }
    });
    this.optionSets.push({
      options: ['memberIds', 'memberUserNames'],
      runsWhen: (args) => {
        return args.options.memberIds !== undefined || args.options.memberUserNames !== undefined;
      }
    });
  }

  #initTypes(): void {
    this.types.boolean.push('isPrivate', 'allowEternalSenders', 'autoSubscribeNewMembers', 'hideFromAddressLists', 'hideFromOutlookClients');
    this.types.string.push('id', 'displayName', 'newDisplayName', 'description', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'logoPath');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.newDisplayName &&
          args.options.description === undefined &&
          args.options.ownerIds === undefined &&
          args.options.ownerUserNames === undefined &&
          args.options.memberIds === undefined &&
          args.options.memberUserNames === undefined &&
          args.options.isPrivate === undefined &&
          args.options.logoPath === undefined &&
          args.options.allowExternalSenders === undefined &&
          args.options.autoSubscribeNewMembers === undefined &&
          args.options.hideFromAddressLists === undefined &&
          args.options.hideFromOutlookClients === undefined) {
          return 'Specify at least one option to update.';
        }

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
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

        if (args.options.logoPath) {
          const fullPath: string = path.resolve(args.options.logoPath);

          if (!fs.existsSync(fullPath)) {
            return `File '${fullPath}' not found`;
          }

          if (fs.lstatSync(fullPath).isDirectory()) {
            return `Path '${fullPath}' points to a directory`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_SET, commands.M365GROUP_SET);

    try {
      if ((args.options.allowExternalSenders !== undefined || args.options.autoSubscribeNewMembers !== undefined) && accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
        throw `Option 'allowExternalSenders' and 'autoSubscribeNewMembers' can only be used when using delegated permissions.`;
      }

      const groupId = args.options.id || await entraGroup.getGroupIdByDisplayName(args.options.displayName!);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
      }

      if (this.verbose) {
        await logger.logToStderr(`Updating Microsoft 365 Group ${args.options.id || args.options.displayName}...`);
      }

      if (args.options.newDisplayName || args.options.description !== undefined || args.options.isPrivate !== undefined) {
        const update: Group = {
          displayName: args.options.newDisplayName,
          description: args.options.description !== '' ? args.options.description : null
        };

        if (args.options.isPrivate !== undefined) {
          update.visibility = args.options.isPrivate ? 'Private' : 'Public';
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${groupId}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: update
        };

        await request.patch(requestOptions);
      }

      // This has to be a separate request due to some Graph API limitations. Otherwise it will throw an error.
      if (args.options.allowExternalSenders !== undefined || args.options.autoSubscribeNewMembers !== undefined || args.options.hideFromAddressLists !== undefined || args.options.hideFromOutlookClients !== undefined) {
        const requestBody: any = {
          allowExternalSenders: args.options.allowExternalSenders,
          autoSubscribeNewMembers: args.options.autoSubscribeNewMembers,
          hideFromAddressLists: args.options.hideFromAddressLists,
          hideFromOutlookClients: args.options.hideFromOutlookClients
        };

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: requestBody
        };
        await request.patch(requestOptions);
      }

      if (args.options.logoPath) {
        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          await logger.logToStderr(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${groupId}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        await this.setGroupLogo(requestOptions, EntraM365GroupSetCommand.numRepeat, logger);
      }
      else if (this.debug) {
        await logger.logToStderr('logoPath not set. Skipping');
      }

      const ownerIds: string[] = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      const memberIds: string[] = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);

      if (ownerIds.length > 0) {
        await this.updateUsers(logger, groupId, 'owners', ownerIds);
      }

      if (memberIds.length > 0) {
        await this.updateUsers(logger, groupId, 'members', memberIds);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async setGroupLogo(requestOptions: any, retryLeft: number, logger: Logger): Promise<void> {
    try {
      await request.put(requestOptions);
    }
    catch (err: any) {
      if (--retryLeft > 0) {
        await setTimeout(this.pollingInterval * (EntraM365GroupSetCommand.numRepeat - retryLeft));
        await this.setGroupLogo(requestOptions, retryLeft, logger);
      }
      else {
        throw err;
      }
    }
  }

  private getImageContentType(imagePath: string): string {
    const extension: string = imagePath.substr(imagePath.lastIndexOf('.')).toLowerCase();

    switch (extension) {
      case '.png':
        return 'image/png';
      case '.gif':
        return 'image/gif';
      default:
        return 'image/jpeg';
    }
  }

  private async getUserIds(logger: Logger, userIds?: string, userNames?: string): Promise<string[]> {
    if (userIds) {
      return formatting.splitAndTrim(userIds);
    }

    if (userNames) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user IDs...`);
      }

      return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
    }

    return [];
  }

  private async updateUsers(logger: Logger, groupId: string, role: 'members' | 'owners', userIds: string[]): Promise<void> {
    const groupUsers = await odata.getAllItems<User>(`${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=id`);
    const userIdsToAdd = userIds.filter(userId => !groupUsers.some(groupUser => groupUser.id === userId));
    const userIdsToRemove = groupUsers.filter(groupUser => !userIds.some(userId => groupUser.id === userId)).map(user => user.id);

    if (this.verbose) {
      await logger.logToStderr(`Adding ${userIdsToAdd.length} ${role}...`);
    }

    for (let i = 0; i < userIdsToAdd.length; i += 400) {
      const userIdsBatch = userIdsToAdd.slice(i, i + 400);
      const batchRequestOptions = this.getBatchRequestOptions();

      // only 20 requests per one batch are allowed
      for (let j = 0; j < userIdsBatch.length; j += 20) {
        // only 20 users can be added in one request
        const userIdsChunk = userIdsBatch.slice(j, j + 20);
        batchRequestOptions.data.requests.push({
          id: j + 1,
          method: 'PATCH',
          url: `/groups/${groupId}`,
          headers: {
            'content-type': 'application/json;odata.metadata=none',
            accept: 'application/json;odata.metadata=none'
          },
          body: {
            [`${role}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
          }
        });
      }

      const res = await request.post<{ responses: { status: number; body: any }[] }>(batchRequestOptions);
      for (const response of res.responses) {
        if (response.status !== 204) {
          throw response.body;
        }
      }
    }

    if (this.verbose) {
      await logger.logToStderr(`Removing ${userIdsToRemove.length} ${role}...`);
    }

    for (let i = 0; i < userIdsToRemove.length; i += 20) {
      const userIdsBatch = userIdsToRemove.slice(i, i + 20);
      const batchRequestOptions = this.getBatchRequestOptions();

      userIdsBatch.map(userId => {
        batchRequestOptions.data.requests.push({
          id: userId,
          method: 'DELETE',
          url: `/groups/${groupId}/${role}/${userId}/$ref`
        });
      });

      const res = await request.post<{ responses: { id: string, status: number; body: any }[] }>(batchRequestOptions);
      for (const response of res.responses) {
        if (response.status !== 204) {
          throw response.body;
        }
      }
    }
  }

  private getBatchRequestOptions(): CliRequestOptions {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/$batch`,
      headers: {
        'content-type': 'application/json;odata.metadata=none',
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        requests: []
      }
    };

    return requestOptions;
  }
}

export default new EntraM365GroupSetCommand();