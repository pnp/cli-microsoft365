import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string().refine(val => validation.isValidTeamsChatId(val), {
    message: 'The value of the option chatId must be a valid Teams ChatId.'
  }).alias('i'),
  userId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value of the option userId must be a valid GUID.'
  }).optional(),
  userName: z.string().refine(val => validation.isValidUserPrincipalName(val), {
    message: 'The value of the option userName must be a valid user principal name.'
  }).optional(),
  role: z.enum(['owner', 'guest']).optional(),
  visibleHistoryStartDateTime: z.string().refine(val => validation.isValidISODateTime(val), {
    message: 'The value of the option visibleHistoryStartDateTime is not a valid ISO date.'
  }).optional(),
  withAllHistory: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a member to a Microsoft Teams chat conversation.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !(opts.userId && opts.userName), {
        message: 'Specify either userId or userName, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      })
      .refine(opts => {
        if (opts.visibleHistoryStartDateTime && opts.withAllHistory) {
          return false;
        }
        return true;
      }, {
        message: 'Specify either visibleHistoryStartDateTime or withAllHistory, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['visibleHistoryStartDateTime', 'withAllHistory']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding member ${args.options.userId || args.options.userName} to chat with id ${args.options.chatId}...`);
      }

      const chatMemberAddOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/chats/${args.options.chatId}/members`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}`,
          visibleHistoryStartDateTime: args.options.withAllHistory ? '0001-01-01T00:00:00Z' : args.options.visibleHistoryStartDateTime,
          roles: [args.options.role || 'owner']
        }
      };

      await request.post(chatMemberAddOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatMemberAddCommand();