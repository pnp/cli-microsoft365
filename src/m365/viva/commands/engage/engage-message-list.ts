import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  olderThanId: z.coerce.number().optional(),
  threaded: z.boolean().optional(),
  limit: z.coerce.number().optional(),
  feedType: z.enum(['All', 'Top', 'My', 'Following', 'Sent', 'Private', 'Received']).optional(),
  groupId: z.coerce.number().optional(),
  threadId: z.coerce.number().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageMessageListCommand extends VivaEngageCommand {
  private items!: any[];

  public get name(): string {
    return commands.ENGAGE_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Returns all accessible messages from the user\'s Viva Engage network';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'replied_to_id', 'thread_id', 'group_id', 'shortBody'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !(opts.groupId && opts.threadId), {
        message: 'You cannot specify groupId and threadId at the same time',
        params: { customCode: 'required' }
      })
      .refine(opts => !(opts.feedType && (opts.groupId || opts.threadId)), {
        message: 'You cannot specify the feedType with groupId or threadId at the same time',
        params: { customCode: 'required' }
      });
  }

  private async getAllItems(logger: Logger, args: CommandArgs, messageId: number): Promise<void> {
    let endpoint = `${this.resource}/v1`;

    if (args.options.threadId) {
      endpoint += `/messages/in_thread/${args.options.threadId}.json`;
    }
    else if (args.options.groupId) {
      endpoint += `/messages/in_group/${args.options.groupId}.json`;
    }
    else {
      if (!args.options.feedType) {
        args.options.feedType = "All";
      }

      switch (args.options.feedType) {
        case 'Top':
          endpoint += `/messages/algo.json`;
          break;
        case 'My':
          endpoint += `/messages/my_feed.json`;
          break;
        case 'Following':
          endpoint += `/messages/following.json`;
          break;
        case 'Sent':
          endpoint += `/messages/sent.json`;
          break;
        case 'Private':
          endpoint += `/messages/private.json`;
          break;
        case 'Received':
          endpoint += `/messages/received.json`;
          break;
        default:
          endpoint += `/messages.json`;
      }
    }

    if (messageId !== -1) {
      endpoint += `?older_than=${messageId}`;
    }
    else if (args.options.olderThanId) {
      endpoint += `?older_than=${args.options.olderThanId}`;
    }

    if (args.options.threaded) {
      if (endpoint.indexOf("?") > -1) {
        endpoint += "&";
      }
      else {
        endpoint += "?";
      }

      endpoint += `threaded=true`;
    }

    const requestOptions: CliRequestOptions = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res: any = await request.get(requestOptions);
    this.items = this.items.concat(res.messages);

    if (args.options.limit && this.items.length > args.options.limit) {
      this.items = this.items.slice(0, args.options.limit);
    }
    else if ((res.meta.older_available === true)) {
      await this.getAllItems(logger, args, this.items[this.items.length - 1].id);
    }
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.items = []; // this will reset the items array in interactive mode

    try {
      await this.getAllItems(logger, args, -1);

      this.items.forEach(m => {
        let shortBody;
        const bodyToProcess = m.body.plain;

        if (bodyToProcess) {
          let maxLength = 35;
          let addedDots = "...";
          if (bodyToProcess.length < maxLength) {
            maxLength = bodyToProcess.length;
            addedDots = "";
          }

          shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
        }

        m.shortBody = shortBody;
      });

      await logger.log(this.items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageMessageListCommand();