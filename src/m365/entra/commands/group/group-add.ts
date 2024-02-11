import { Group } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description?: string;
  type: string;
  mailNickname?: string;
  ownerIds?: string;
  ownerUserNames?: string;
  memberIds?: string;
  memberUserNames?: string;
  visibility?: string;
}

class EntraGroupAddCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Entra group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUP_ADD];
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-t, --type <type>',
        autocomplete: ['microsoft365', 'security']
      },
      {
        option: '-m, --mailNickname [mailNickname]'
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
        autocomplete: ['Public', 'Private', 'HiddenMembership']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.displayName.length > 256) {
          return `The maximum amount of characters for 'displayName' is 256.`;
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

        if (['microsoft365', 'security'].indexOf(args.options.type) === -1) {
          return `Option 'type' must be one of the following values: microsoft365, security.`;
        }

        if (args.options.type === 'microsoft365' && !args.options.visibility) {
          return `Option 'visibility' must be specified if the option 'type' is set to microsoft365`;
        }

        if (args.options.visibility && ['Public', 'Private', 'HiddenMembership'].indexOf(args.options.visibility!) === -1) {
          return `Option 'visibility' must be one of the following values: Public, Private, HiddenMembership.`;
        }

        return true;
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        type: typeof args.options.type !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
        ownerIds: typeof args.options.ownerIds !== 'undefined',
        ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
        memberIds: typeof args.options.memberIds !== 'undefined',
        memberUserNames: typeof args.options.memberUserNames !== 'undefined',
        visibility: typeof args.options.visibility !== 'undefined'
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: Group;
    let ownerIds: string[] = [];
    let memberIds: string[] = [];

    try {
      const manifest = this.createRequestBody(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: manifest
      };

      ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);

      group = await request.post<Group>(requestOptions);

      if (ownerIds.length !== 0) {
        await this.addUsers(group.id!, 'owners', ownerIds);
      }

      if (memberIds.length !== 0) {
        await this.addUsers(group.id!, 'members', memberIds);
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  };

  private createRequestBody(options: Options): any {
    const requestBody: any = {
      displayName: options.displayName,
      description: options.description,
      mailNickName: options.mailNickname ?? this.generateMailNickname(),
      visibility: options.visibility ?? 'Public',
      groupTypes: options.type === 'microsoft365' ? ['Unified'] : [],
      mailEnabled: options.type === 'security' ? false : true,
      securityEnabled: true
    };

    this.addUnknownOptionsToPayload(requestBody, options);
    return requestBody;
  }

  private generateMailNickname(): string {
    return `Group${Math.floor(Math.random() * 1000000)}`;
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

export default new EntraGroupAddCommand();