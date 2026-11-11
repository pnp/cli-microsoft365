import { Conversation } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  groupId: z.guid().optional().alias('i'),
  groupName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupConversationListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_LIST;
  }

  public get description(): string {
    return 'Lists conversations for the specified Microsoft 365 group';
  }

  public defaultProperties(): string[] | undefined {
    return ['topic', 'lastDeliveredDateTime', 'id'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.groupId, opts.groupName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: groupId or groupName.',
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving conversations for Microsoft 365 Group: ${args.options.groupId || args.options.groupName}...`);
    }

    try {
      let groupId = args.options.groupId;

      if (args.options.groupName) {
        groupId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
      }

      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId!);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
      }

      const conversations = await odata.getAllItems<Conversation>(`${this.resource}/v1.0/groups/${groupId}/conversations`);
      await logger.log(conversations);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupConversationListCommand();