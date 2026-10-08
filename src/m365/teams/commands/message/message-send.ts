import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  message: z.string()
    .alias('m')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMessageSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_SEND;
  }

  public get description(): string {
    return 'Sends a message to a channel in a Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          body: {
            contentType: 'html',
            content: args.options.message
          }
        }
      };

      const response = await request.post<any>(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMessageSendCommand();