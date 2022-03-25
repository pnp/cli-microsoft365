import { AadUserConversationMember, Chat, ConversationMember } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import * as os from 'os';
import Auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken, odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId?: string;
  userEmails?: string;
  chatName?: string;
  message: string;
}

class TeamsChatMessageSendCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MESSAGE_SEND;
  }

  public get description(): string {
    return 'Sends a chat message to a Microsoft Teams chat conversation.';
  }

  public getTelemetryProperties(args: any): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.chatId = typeof args.options.chatId !== 'undefined';
    telemetryProps.userEmails = typeof args.options.userEmails !== 'undefined';
    telemetryProps.chatName = typeof args.options.chatName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {    
    this
      .getChatId(logger, args)
      .then(chatId => this.sendChatMessage(chatId as string, args))
      .then(_ => {
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));   
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--chatId [chatId]'
      },
      {
        option: '-e, --userEmails [userEmails]'
      },
      {
        option: '--chatName [chatName]'
      },
      {
        option: '-m, --message <message>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.chatId && !args.options.userEmails && !args.options.chatName) {
      return 'Specify chatId or userEmails or chatName, one is required.';
    }

    let nrOfMutuallyExclusiveOptionsInUse = 0;
    if (args.options.chatId) { nrOfMutuallyExclusiveOptionsInUse++; }
    if (args.options.userEmails) { nrOfMutuallyExclusiveOptionsInUse++; }
    if (args.options.chatName) { nrOfMutuallyExclusiveOptionsInUse++; }

    if (nrOfMutuallyExclusiveOptionsInUse > 1) {
      return 'Specify either chatId or userEmails or chatName, but not multiple.';
    }

    if (!args.options.message) {
      return 'Specify a message to send.';
    }

    if (args.options.chatId && !validation.isValidTeamsChatId(args.options.chatId)) {
      return `${args.options.chatId} is not a valid Teams ChatId.`;
    }

    if (args.options.userEmails) {
      const userEmails = this.convertCommaSeparatedListToArray(args.options.userEmails);
      if (!userEmails || userEmails.length === 0 || userEmails.some(e => !validation.isValidUserPrincipalName(e))) {
        return `${args.options.userEmails} contains one or more invalid email addresses.`;
      }
    }

    return true;
  }

  private async getChatId(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.chatId) {
      return args.options.chatId;
    }    

    return args.options.userEmails 
      ? this.ensureChatIdByUserEmails(args.options.userEmails, logger)
      : this.getChatIdByName(args.options.chatName as string, logger);
  }

  private async ensureChatIdByUserEmails(userEmailsOption: string, logger: Logger): Promise<string> {
    const userEmails = this.convertCommaSeparatedListToArray(userEmailsOption);
    const currentUserEmail = accessToken.getUserNameFromAccessToken(Auth.service.accessTokens[this.resource].accessToken).toLowerCase();
    const existingChats = await this.findExistingGroupChatsByMembers([currentUserEmail, ...userEmails], logger);
    
    if (!existingChats || existingChats.length === 0) {
      const chat = await this.createConversation([currentUserEmail, ...userEmails]);
      return chat.id as string;
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const disambiguationText = existingChats.map(c => {
      return `- ${c.id}${c.topic && ' - '}${c.topic} - ${c.createdDateTime && new Date(c.createdDateTime).toLocaleString()}`;
    }).join(os.EOL);

    throw new Error(`Multiple chat conversations with this name found. Please disambiguate:${os.EOL}${disambiguationText}`);
  }

  private async getChatIdByName(chatName: string, logger: Logger): Promise<string> {
    const existingChats = await this.findExistingGroupChatsByName(chatName, logger);

    if (!existingChats || existingChats.length === 0) {
      throw new Error('No chat conversation was found with this name.');
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const disambiguationText = existingChats.map(c => {
      const memberstring = (c.members as ConversationMember[]).map(m => (m as AadUserConversationMember).email).join(', ');
      return `- ${c.id} - ${c.createdDateTime && new Date(c.createdDateTime).toLocaleString()} - ${memberstring}`;
    }).join(os.EOL);

    throw new Error(`Multiple chat conversations with this name found. Please disambiguate:${os.EOL}${disambiguationText}`);
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

      const requestOptions: AxiosRequestConfig = {
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
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/chats/${chatId}/messages`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        body: {
          content: args.options.message
        }
      }
    };

    await request.post(requestOptions);
  }

  private async findExistingGroupChatsByMembers(expectedMemberEmails: string[], logger: Logger): Promise<Chat[]> {
    const endpoint = `${this.resource}/v1.0/chats?$filter=chatType eq 'group'&$expand=members&$select=id,topic,createdDateTime,members`;
    const foundChats: Chat[] = [];

    const chats = await odata.getAllItems<Chat>(endpoint, logger);

    for (const chat of chats) {
      const chatMembers = chat.members as ConversationMember[];
      if (chatMembers.length === expectedMemberEmails.length) {
        const chatMemberEmails = chatMembers.map(member => (member as AadUserConversationMember).email?.toLowerCase());

        if (expectedMemberEmails.every(email => chatMemberEmails.some(memberEmail => memberEmail === email))) {
          foundChats.push(chat);
        }
      }
    }

    return foundChats;
  }

  private async findExistingGroupChatsByName(name: string, logger: Logger): Promise<Chat[]> {
    const endpoint = `${this.resource}/v1.0/chats?$filter=topic eq '${encodeURIComponent(name)}'&$expand=members&$select=id,topic,createdDateTime,chatType`;
    return odata.getAllItems<Chat>(endpoint, logger);
  }

  private convertCommaSeparatedListToArray(userEmailsString: string): string[] {
    return userEmailsString.toLowerCase().replace(/\s/g, '').split(',').filter(e => e && e !== '');
  }
}

module.exports = new TeamsChatMessageSendCommand();