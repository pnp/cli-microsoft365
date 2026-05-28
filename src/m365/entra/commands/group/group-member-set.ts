import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const RoleEnum = {
  Owner: 'Owner',
  Member: 'Member'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  groupId: z.uuid().optional().alias('i'),
  groupName: z.string().optional().alias('n'),
  userIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'userIds': ${e.input}.`
    }).optional(),
  userNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `User principal name '${e.input}' is invalid for option 'userNames'.`
    }).optional(),
  role: zod.coercedEnum(RoleEnum).alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupMemberSetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_SET;
  }

  public get description(): string {
    return 'Updates the role of members in a Microsoft Entra ID group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.groupId, options.groupName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: groupId or groupName.',
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName']
        }
      })
      .refine(options => [options.userIds, options.userNames].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: userIds or userNames.',
        params: {
          customCode: 'optionSet',
          options: ['userIds', 'userNames']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding member(s) ${args.options.userIds || args.options.userNames} to role ${args.options.role} of group ${args.options.groupId || args.options.groupName}...`);
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
      await logger.logToStderr(`Retrieving ID of group ${options.groupName}...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupName!);
  }

  private async getUserIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.userIds) {
      return options.userIds.split(',').map(i => i.trim());
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
      await logger.logToStderr(`Removing members from the old role '${currentRole}'.`);
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

export default new EntraGroupMemberSetCommand();