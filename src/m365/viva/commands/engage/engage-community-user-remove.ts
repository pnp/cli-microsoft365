import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    communityId: z.string().optional(),
    communityDisplayName: zod.alias('n', z.string().optional()),
    entraGroupId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    id: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    userName: z.string()
      .refine(userName => validation.isValidUserPrincipalName(userName), userName => ({
        message: `'${userName}' is not a valid user principal name.`
      })).optional(),
    force: z.boolean().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunityUserRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specified user from a Microsoft 365 Viva Engage community';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.communityId, options.communityDisplayName, options.entraGroupId].filter(x => x !== undefined).length === 1, {
        message: 'Specify either communityId, communityDisplayName, or entraGroupId, but not multiple.'
      })
      .refine(options => options.communityId || options.communityDisplayName || options.entraGroupId, {
        message: 'Specify at least one of communityId, communityDisplayName, or entraGroupId.'
      })
      .refine(options => options.id || options.userName, {
        message: 'Specify either of id or userName.'
      })
      .refine(options => options.userName !== undefined || options.id !== undefined, {
        message: 'Specify either id or userName, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        await this.deleteUserFromCommunity(args.options, logger);
      }
      else {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the user ${args.options.id || args.options.userName} from the community ${args.options.communityDisplayName || args.options.communityId || args.options.entraGroupId}?` });

        if (result) {
          await this.deleteUserFromCommunity(args.options, logger);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async deleteUserFromCommunity(options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing user from community...');
    }

    let entraGroupId = options.entraGroupId;

    if (options.communityDisplayName) {
      const community = await vivaEngage.getCommunityByDisplayName(options.communityDisplayName, ['groupId']);
      entraGroupId = community.groupId;
    }
    else if (options.communityId) {
      const community = await vivaEngage.getCommunityById(options.communityId, ['groupId']);
      entraGroupId = community.groupId;
    }

    const userId = options.id || await entraUser.getUserIdByUpn(options.userName!);

    await this.deleteUser(entraGroupId!, userId, 'owners');
    await this.deleteUser(entraGroupId!, userId, 'members');
  }

  private async deleteUser(entraGroupId: string, userId: string, role: string): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${entraGroupId}/${role}/${userId}/$ref`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      await request.delete(requestOptions);
    }
    catch (err: any) {
      if (err.response.status !== 404) {
        throw err.response.data;
      }
    }
  }
}

export default new VivaEngageCommunityUserRemoveCommand();