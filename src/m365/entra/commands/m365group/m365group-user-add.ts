import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const RoleEnum = {
  owner: 'owner',
  member: 'member'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  ids: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'ids': ${validation.isValidGuidArray(e.input as string)}.`
    }).optional(),
  userNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for the option 'userNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).optional(),
  groupId: z.guid().optional().alias('i'),
  groupName: z.string().optional(),
  teamId: z.guid().optional(),
  teamName: z.string().optional(),
  role: zod.coercedEnum(RoleEnum).optional().alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupUserAddCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Adds user to specified Microsoft 365 Group or Microsoft Teams team';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.groupId, opts.groupName, opts.teamId, opts.teamName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: groupId, groupName, teamId or teamName.',
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName', 'teamId', 'teamName']
        }
      })
      .refine(opts => [opts.ids, opts.userNames].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: ids or userNames.',
        params: {
          customCode: 'optionSet',
          options: ['ids', 'userNames']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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
      return formatting.splitAndTrim(userIds);
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving user ID(s) by username(s)...');
    }

    return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames!));
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