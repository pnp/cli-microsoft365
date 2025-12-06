import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { entraUser } from '../../../../utils/entraUser.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    ids: zod.alias('ids', z.string().optional()),
    userNames: zod.alias('userNames', z.string().optional()),
    groupId: zod.alias('i', z.string().uuid().optional()),
    groupName: zod.alias('groupName', z.string().optional()),
    teamId: zod.alias('teamId', z.string().uuid().optional()),
    teamName: zod.alias('teamName', z.string().optional()),
    role: zod.alias('r', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupUserAddCommand extends GraphCommand {
  private readonly allowedRoles: string[] = ['owner', 'member'];

  public get name(): string {
    return commands.M365GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Adds user to specified Microsoft 365 Group or Microsoft Teams team';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public alias(): string[] | undefined {
    const teamCommands: string[] = [
      'teams user add'
    ];

    return teamCommands;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.groupId, options.groupName, options.teamId, options.teamName].filter(Boolean).length === 1, {
        message: 'Specify either groupId, groupName, teamId, or teamName'
      })
      .refine(options => [options.ids, options.userNames].filter(Boolean).length === 1, {
        message: 'Specify either ids or userNames'
      })
      .refine(options => {
        if (options.ids) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(options.ids);
          return isValidGUIDArrayResult === true;
        }
        return true;
      }, {
        message: 'The following GUIDs are invalid for the option \'ids\''
      })
      .refine(options => {
        if (options.userNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(options.userNames);
          return isValidUPNArrayResult === true;
        }
        return true;
      }, {
        message: 'The following user principal names are invalid for the option \'userNames\''
      })
      .refine(options => {
        if (options.role) {
          return this.allowedRoles.some(role => role.toLowerCase() === options.role!.toLowerCase());
        }
        return true;
      }, {
        message: 'Invalid role value. Allowed values are: owner,member'
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