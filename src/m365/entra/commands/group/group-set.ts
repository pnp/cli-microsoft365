import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

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
  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Entra group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUP_SET];
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
        autocomplete: ['Public', 'Private']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.newDisplayName && args.options.newDisplayName.length > 256) {
          return `The maximum amount of characters for 'newDisplayName' is 256.`;
        }

        if (args.options.mailNickname) {
          if (!validation.isValidMailNickname(args.options.mailNickname)) {
            return `Value for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \ [] " ; : <> , SPACE.`;
          }

          if (args.options.mailNickname.length > 64) {
            return `The maximum amount of characters for 'mailNickname' is 64.`;
          }
        }

        if (args.options.ownerIds) {
          const ids = args.options.ownerIds.split(',').map(i => i.trim());
          if (!validation.isValidGuidArray(ids)) {
            const invalidGuid = ids.find(id => !validation.isValidGuid(id));
            return `'${invalidGuid}' is not a valid GUID for option 'ownerIds'.`;
          }
        }

        if (args.options.ownerUserNames) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `User principal name '${isValidUserPrincipalNameArray}' is invalid for option 'ownerUserNames'.`;
          }
        }

        if (args.options.memberIds) {
          const ids = args.options.memberIds.split(',').map(i => i.trim());
          if (!validation.isValidGuidArray(ids)) {
            const invalidGuid = ids.find(id => !validation.isValidGuid(id));
            return `'${invalidGuid}' is not a valid GUID for option 'memberIds'.`;
          }
        }

        if (args.options.memberUserNames) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.memberUserNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `User principal name '${isValidUserPrincipalNameArray}' is invalid for option 'memberUserNames'.`;
          }
        }

        if (args.options.visibility && ['Public', 'Private'].indexOf(args.options.visibility!) === -1) {
          return `Option 'visibility' must be one of the following values: Public, Private.`;
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

      const manifest = this.createRequestBody(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${groupId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: manifest
      };

      ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);

      await request.patch(requestOptions);

      if (ownerIds.length !== 0) {
        await this.addUsers(groupId!, 'owners', ownerIds);
      }

      if (memberIds.length !== 0) {
        await this.addUsers(groupId!, 'members', memberIds);
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

  private async getUserIds(logger: Logger, userIds: string | undefined, userNames: string | undefined): Promise<string[]> {
    if (userIds) {
      return userIds.split(',').map(o => o.trim());
    }

    if (!userNames) {
      if (this.verbose) {
        await logger.logToStderr('No users to validate, skipping.');
      }
      return [];
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving user information.');
    }

    const userArr: string[] = userNames.split(',').map(o => o.trim());

    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of user(s)...');
    }

    return entraUser.getUserIdsByUpns(userArr);
  }

  private async addUsers(groupId: string, role: string, userIds: string[]): Promise<void> {
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
  }
}

export default new EntraGroupSetCommand();