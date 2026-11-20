import { z } from 'zod';
import { Chat } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { chatUtil } from './chatUtil.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string().refine(val => validation.isValidTeamsChatId(val), {
    message: 'The value is not a valid Teams ChatId.'
  }).optional(),
  userEmails: z.string().refine(val => {
    const userEmails = val.trim().toLowerCase().split(',').filter(e => e && e !== '');
    if (!userEmails || userEmails.length === 0) {
      return false;
    }
    return userEmails.every(e => validation.isValidUserPrincipalName(e));
  }, {
    message: 'The option userEmails contains one or more invalid email addresses.'
  }).alias('e').optional(),
  chatName: z.string().optional(),
  message: z.string().alias('m'),
  contentType: z.enum(['text', 'html']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatMessageSendCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.CHAT_MESSAGE_SEND;
  }

  public get description(): string {
    return 'Sends a chat message to a Microsoft Teams chat conversation.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.chatId, opts.userEmails, opts.chatName].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of chatId, userEmails or chatName, but not more than one.',
        params: {
          customCode: 'optionSet',
          options: ['chatId', 'userEmails', 'chatName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const chatId = await this.getChatId(logger, args);
      await this.sendChatMessage(chatId, args);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getChatId(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.chatId) {
      return args.options.chatId;
    }

    return args.options.userEmails
      ? this.ensureChatIdByUserEmails(args.options.userEmails)
      : this.getChatIdByName(args.options.chatName as string);
  }

  private async ensureChatIdByUserEmails(userEmailsOption: string): Promise<string> {
    const userEmails = userEmailsOption.trim().toLowerCase().split(',').filter(e => e && e !== '');
    const currentUserEmail = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken).toLowerCase();
    const existingChats = await chatUtil.findExistingChatsByParticipants([currentUserEmail, ...userEmails]);

    if (!existingChats || existingChats.length === 0) {
      const chat = await this.createConversation([currentUserEmail, ...userEmails]);
      return chat.id as string;
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
    const result = await cli.handleMultipleResultsFound<Chat>(`Multiple chat conversations with this name found.`, resultAsKeyValuePair);
    return result.id!;
  }

  private async getChatIdByName(chatName: string): Promise<string> {
    const existingChats = await chatUtil.findExistingGroupChatsByName(chatName);

    if (!existingChats || existingChats.length === 0) {
      throw 'No chat conversation was found with this name.';
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
    const result = await cli.handleMultipleResultsFound<Chat>(`Multiple chat conversations with this name found.`, resultAsKeyValuePair);
    return result.id!;
  }

  // This Microsoft Graph API request throws an intermittent 404 exception, saying that it cannot find the principal.
  // The same behavior occurs when creating the conversation through the Graph Explorer.
  // It seems to happen when the userEmail casing does not match the casing of the actual UPN. 
  // When the first request throws an error, the second request does succeed. 
  // Therefore a retry-mechanism is implemented here. 
  private async createConversation(memberEmails: string[], retried: number = 0): Promise<Chat> {
    try {
      const jsonBody = {
        chatType: memberEmails.length > 2 ? 'group' : 'oneOnOne',
        members: memberEmails.map(email => {
          return {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${email}`
          };
        })
      };

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/chats`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: jsonBody
      };

      return await request.post<Chat>(requestOptions);
    }
    catch (err) {
      if ((err as Error).message?.indexOf('404') > -1 && retried < 4) {
        return await this.createConversation(memberEmails, retried + 1);
      }

      throw err;
    }
  }

  private async sendChatMessage(chatId: string, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/chats/${chatId}/messages`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        body: {
          contentType: args.options.contentType || 'text',
          content: args.options.message
        }
      }
    };

    return request.post(requestOptions);
  }
}

export default new TeamsChatMessageSendCommand();