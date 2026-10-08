import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  messageId: z.coerce.number(),
  enable: z.boolean().optional(),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageMessageLikeSetCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_MESSAGE_LIKE_SET;
  }

  public get description(): string {
    return 'Likes or unlikes a Viva Engage message';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.enable === false) {
      if (args.options.force) {
        await this.executeLikeAction(args.options);
      }
      else {
        const message = `Are you sure you want to unlike message ${args.options.messageId}?`;

        const result = await cli.promptForConfirmation({ message });

        if (result) {
          await this.executeLikeAction(args.options);
        }
      }
    }
    else {
      await this.executeLikeAction(args.options);
    }
  }

  private async executeLikeAction(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1/messages/liked_by/current.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        message_id: options.messageId
      }
    };

    try {
      if (options.enable !== false) {
        await request.post(requestOptions);
      }
      else {
        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageMessageLikeSetCommand();