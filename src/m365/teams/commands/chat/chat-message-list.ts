import { ItemBody, Message } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

interface ExtendedMessage extends Message {
  shortBody?: string;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string()
    .refine(id => validation.isValidTeamsChatId(id), {
      error: e => `'${e.input}' is not a valid value for option chatId.`
    })
    .alias('i'),
  endDateTime: z.string()
    .refine(time => validation.isValidISODateTime(time), {
      error: e => `'${e.input}' is not a valid ISO date-time string for option endDateTime.`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsChatMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Lists all messages from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'createdDateTime', 'shortBody'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let apiUrl = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;

      if (args.options.endDateTime) {
        // You can only filter results if the request URL contains the $orderby and $filter query parameters configured for the same property;
        // otherwise, the $filter query option is ignored.
        apiUrl += `?$filter=createdDateTime lt ${args.options.endDateTime}&$orderby=createdDateTime desc`;
      }

      const items = await odata.getAllItems<ExtendedMessage>(apiUrl);
      if (args.options.output && args.options.output !== 'json') {
        items.forEach(i => {
          // hoist the content to body for readability
          i.body = (i.body as ItemBody).content as any;

          let shortBody: string | undefined;
          const bodyToProcess = i.body as string;

          if (bodyToProcess) {
            let maxLength = 50;
            let addedDots = '...';
            if (bodyToProcess.length < maxLength) {
              maxLength = bodyToProcess.length;
              addedDots = '';
            }

            shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
          }

          i.shortBody = shortBody;
        });
      }

      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatMessageListCommand();