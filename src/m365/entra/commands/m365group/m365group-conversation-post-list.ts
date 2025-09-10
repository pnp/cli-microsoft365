import { Post } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    groupId: zod.alias('i', z.string().uuid().optional()),
    groupName: zod.alias('d', z.string().optional()),
    threadId: zod.alias('t', z.string())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupConversationPostListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_POST_LIST;
  }

  public get description(): string {
    return 'Lists conversation posts of a Microsoft 365 group';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.groupId, options.groupName].filter(Boolean).length === 1, {
        message: 'Specify either groupId or groupName'
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['receivedDateTime', 'id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const retrievedgroupId = await this.getGroupId(args);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(retrievedgroupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${retrievedgroupId}' is not a Microsoft 365 group.`);
      }

      const posts = await odata.getAllItems<Post>(`${this.resource}/v1.0/groups/${retrievedgroupId}/threads/${args.options.threadId}/posts`);
      await logger.log(posts);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.groupId) {
      return formatting.encodeQueryParameter(args.options.groupId);
    }

    const group = await entraGroup.getGroupByDisplayName(args.options.groupName!);
    return group.id!;
  }
}

export default new EntraM365GroupConversationPostListCommand();