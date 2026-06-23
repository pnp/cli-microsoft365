import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import { User } from '@microsoft/microsoft-graph-types';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const VisibilityEnum = {
  Public: 'Public',
  Private: 'Private'
} as const;

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n'),
  newDisplayName: z.string().max(256, `The maximum amount of characters for 'newDisplayName' is 256.`).optional(),
  description: z.string().optional(),
  mailNickname: z.string()
    .refine(val => validation.isValidMailNickname(val), {
      error: e => `Value '${e.input}' for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \\ [] " ; : <> , SPACE.`
    })
    .refine(val => val.length <= 64, {
      error: `The maximum amount of characters for 'mailNickname' is 64.`
    })
    .optional(),
  ownerIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for option 'ownerIds': ${validation.isValidGuidArray(e.input as string)}.`
    }).optional(),
  ownerUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for option 'ownerUserNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).optional(),
  memberIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for option 'memberIds': ${validation.isValidGuidArray(e.input as string)}.`
    }).optional(),
  memberUserNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for option 'memberUserNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).optional(),
  visibility: zod.coercedEnum(VisibilityEnum).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupSetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Entra group';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.displayName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id or displayName.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName']
        }
      })
      .refine(options => !(options.ownerIds && options.ownerUserNames), {
        error: 'Use one of the following options: ownerIds or ownerUserNames.',
        params: {
          customCode: 'optionSet',
          options: ['ownerIds', 'ownerUserNames']
        }
      })
      .refine(options => !(options.memberIds && options.memberUserNames), {
        error: 'Use one of the following options: memberIds or memberUserNames.',
        params: {
          customCode: 'optionSet',
          options: ['memberIds', 'memberUserNames']
        }
      })
      .refine(options => options.newDisplayName !== undefined || options.description !== undefined || options.visibility !== undefined
        || options.ownerIds !== undefined || options.ownerUserNames !== undefined || options.memberIds !== undefined
        || options.memberUserNames !== undefined || options.mailNickname !== undefined, {
        error: `Specify at least one of the following options: 'newDisplayName', 'description', 'visibility', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'mailNickname'.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let groupId = args.options.id;

    try {
      if (args.options.displayName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving group id...`);
        }

        groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
      }

      const requestBody = {
        displayName: args.options.newDisplayName,
        description: args.options.description === '' ? null : args.options.description,
        mailNickName: args.options.mailNickname,
        visibility: args.options.visibility
      };

      this.addUnknownOptionsToPayloadZod(requestBody, args.options);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${groupId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: requestBody
      };

      await request.patch(requestOptions);

      const ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
      if (ownerIds.length !== 0) {
        await this.updateUsers(logger, groupId!, 'owners', ownerIds);
      }
      else if (this.verbose) {
        await logger.logToStderr(`No owners to update.`);
      }

      const memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);
      if (memberIds.length !== 0) {
        await this.updateUsers(logger, groupId!, 'members', memberIds);
      }
      else if (this.verbose) {
        await logger.logToStderr(`No members to update.`);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  };

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

export default new EntraGroupSetCommand();