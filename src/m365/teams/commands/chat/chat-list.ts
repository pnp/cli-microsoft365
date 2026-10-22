import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(['oneOnOne', 'group', 'meeting']).alias('t').optional(),
  userId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value of the option userId must be a valid GUID.'
  }).optional(),
  userName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_LIST;
  }

  public get description(): string {
    return 'Lists all chat conversations';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'topic', 'chatType'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !(opts.userId && opts.userName), {
        message: 'You can only specify either userId or userName.',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);

    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      throw `The option 'userId' or 'userName' is required when obtaining chats using app only permissions`;
    }
    else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
      throw `The options 'userId' or 'userName' cannot be used when obtaining chats using delegated permissions`;
    }

    let requestUrl = `${this.resource}/v1.0/${!isAppOnlyAccessToken ? 'me' : `users/${args.options.userId || args.options.userName}`}/chats`;

    if (args.options.type) {
      requestUrl += `?$filter=chatType eq '${args.options.type}'`;
    }

    try {
      const items = await odata.getAllItems(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatListCommand();