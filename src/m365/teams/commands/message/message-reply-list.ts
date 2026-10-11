import { z } from 'zod';
import { ChatMessage } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .alias('i'),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams ChannelId.'
    })
    .alias('c'),
  messageId: z.string()
    .alias('m')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMessageReplyListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_REPLY_LIST;
  }

  public get description(): string {
    return 'Retrieves replies to a message from a channel in a Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'body'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}/replies`;

    try {
      const items = await odata.getAllItems<ChatMessage>(endpoint);
      if (args.options.output !== 'json') {
        items.forEach(i => {
          i.body = i.body!.content as any;
        });
      }

      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMessageReplyListCommand();