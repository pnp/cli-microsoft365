import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import teamsCommands from '../../../teams/commands.js';

const options = globalOptionsZod
  .extend({
    teamId: zod.alias('teamId', z.string().uuid().optional()),
    teamName: zod.alias('teamName', z.string().optional()),
    groupId: zod.alias('i', z.string().uuid().optional()),
    groupName: zod.alias('groupName', z.string().optional()),
    ids: zod.alias('ids', z.string().optional()),
    userNames: zod.alias('userNames', z.string().optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupUserRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public alias(): string[] | undefined {
    return [teamsCommands.USER_REMOVE];
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !!(options.groupId || options.groupName || options.teamId || options.teamName), {
        message: 'Specify either groupId, groupName, teamId, or teamName'
      })
      .refine(options => !!(options.ids || options.userNames), {
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeUser = async (): Promise<void> => {
      try {
        const groupId: string = await this.getGroupId(logger, args.options.groupId, args.options.teamId, args.options.groupName, args.options.teamName);
        const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

        if (!isUnifiedGroup) {
          throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
        }

        const userNames = args.options.userNames;
        const userIds: string[] = await this.getUserIds(logger, args.options.ids, userNames);

        await this.removeUsersFromGroup(groupId, userIds, 'owners');
        await this.removeUsersFromGroup(groupId, userIds, 'members');
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeUser();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${args.options.userNames || args.options.ids} from ${args.options.groupId || args.options.groupName || args.options.teamId || args.options.teamName}?` });

      if (result) {
        await removeUser();
      }
    }
  }

  private async getGroupId(logger: Logger, groupId?: string, teamId?: string, groupName?: string, teamName?: string): Promise<string> {
    const id = groupId || teamId;
    if (id) {
      return id;
    }

    const name = groupName ?? teamName;
    if (this.verbose) {
      await logger.logToStderr(`Retrieving Group ID by display name ${name}...`);
    }

    return entraGroup.getGroupIdByDisplayName(name!);
  }

  private async getUserIds(logger: Logger, userIds?: string, userNames?: string): Promise<string[]> {
    if (userIds) {
      return formatting.splitAndTrim(userIds);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving user IDs for {userNames}...`);
    }

    return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames!));
  }

  private async removeUsersFromGroup(groupId: string, userIds: string[], role: string): Promise<void> {
    for (const userId of userIds) {
      try {
        await request.delete({
          url: `${this.resource}/v1.0/groups/${groupId}/${role}/${userId}/$ref`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        });
      }
      catch (err: any) {
        // the 404 error is accepted
        if (err.response.status !== 404) {
          throw err.response.data;
        }
      }
    }
  }
}

export default new EntraM365GroupUserRemoveCommand();