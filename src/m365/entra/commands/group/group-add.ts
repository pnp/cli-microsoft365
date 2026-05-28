import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const GroupTypeEnum = {
  microsoft365: 'microsoft365',
  security: 'security'
} as const;

const VisibilityEnum = {
  Public: 'Public',
  Private: 'Private',
  HiddenMembership: 'HiddenMembership'
} as const;

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  displayName: z.string().max(256, `The maximum amount of characters for 'displayName' is 256.`).alias('n'),
  description: z.string().optional().alias('d'),
  type: zod.coercedEnum(GroupTypeEnum).alias('t'),
  mailNickname: z.string()
    .refine(val => validation.isValidMailNickname(val), {
      error: `Value for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \\ [] " ; : <> , SPACE.`
    })
    .refine(val => val.length <= 64, {
      error: `The maximum amount of characters for 'mailNickname' is 64.`
    })
    .optional().alias('m'),
  ownerIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'ownerIds': ${e.input}.`
    }).optional(),
  ownerUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for the option 'ownerUserNames': ${e.input}.`
    }).optional(),
  memberIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'memberIds': ${e.input}.`
    }).optional(),
  memberUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for the option 'memberUserNames': ${e.input}.`
    }).optional(),
  visibility: zod.coercedEnum(VisibilityEnum).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupAddCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Entra group';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.type !== 'microsoft365' || options.visibility !== undefined, {
        error: `Option 'visibility' must be specified if the option 'type' is set to microsoft365`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: Group;

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

      const ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      const memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);

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

    this.addUnknownOptionsToPayloadZod(requestBody, options);

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