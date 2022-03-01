import { AadUserConversationMember, Chat, ConversationMember } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import Auth from '../../../../Auth';
import * as os from 'os';
import request from '../../../../request';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';
import { ODataResponse } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  participants?: string;
  name?: string;
}

class TeamsChatGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_GET;
  }

  public get description(): string {
    return 'Get a chat conversations by id, participants or chat name';
  }

  public async commandAction(logger: Logger, args: CommandArgs, cb: () => void): Promise<void> {

    try {
      const chatId = args.options.id
        || args.options.participants && await this.getChatIdByParticipants(args.options.participants)
        || args.options.name && await this.getChatIdByName(args.options.name);

      const chat = await this.getChatDetailsById(chatId as string);

      logger.log(chat);
      cb();
    }
    catch(error) {
      this.handleRejectedODataJsonPromise(error, logger, cb);
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-p, --participants [participants]'
      },
      {
        option: '-n, --name [name]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id && !args.options.participants && !args.options.name) {
      return 'Specify id or participants or name, one is required.';
    }

    let nrOfMutuallyExclusiveOptionsInUse = 0;
    if (args.options.id) { nrOfMutuallyExclusiveOptionsInUse++; }
    if (args.options.participants) { nrOfMutuallyExclusiveOptionsInUse++; }
    if (args.options.name) { nrOfMutuallyExclusiveOptionsInUse++; }

    if (nrOfMutuallyExclusiveOptionsInUse > 1) {
      return 'Specify either id or participants or name, but not multiple.';
    }

    if (args.options.id && !validation.isValidTeamsChatId(args.options.id)) {
      return `${args.options.id} is not a valid Teams ChatId.`;
    }

    if (args.options.participants) {
      const participants = args.options.participants.toLowerCase().replace(/\s/g, '').split(',').filter(e => e && e !== '');
      if (!participants || participants.length === 0 || participants.some(e => !validation.isValidUserPrincipalName(e))) {
        return `${args.options.participants} contains one or more invalid email addresses.`;
      }
    }

    return true;
  }
  
  private async getChatDetailsById(id: string): Promise<Chat> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/chats/${encodeURIComponent(id)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'      
    };

    const chat = await request.get<Chat>(requestOptions);
    return chat;
  }

  private async getChatIdByParticipants(participantsString: string): Promise<string> {
    const participants = participantsString.toLowerCase().replace(/\s/g, '').split(',').filter(e => e && e !== '');
    const currentUserEmail = accessToken.getUserNameFromAccessToken(Auth.service.accessTokens[this.resource].accessToken).toLowerCase();
    const existingChats = await this.findExistingChatsByParticipants([currentUserEmail, ...participants]);
    
    if (!existingChats || existingChats.length === 0) {
      throw new Error('No chat conversation was found with these participants.');
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const disambiguationText = existingChats.map(c => {
      return `- ${c.id}${c.topic && ' - '}${c.topic} - ${c.createdDateTime && new Date(c.createdDateTime).toLocaleString()}`;
    }).join(os.EOL);

    throw new Error(`Multiple chat conversations with this topic found. Please disambiguate:${os.EOL}${disambiguationText}`);
    
  }
  
  private async getChatIdByName(name: string): Promise<string> {
    const existingChats = await this.findExistingGroupChatsByTopic(name);

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

    throw new Error(`Multiple chat conversations with this topic found. Please disambiguate:${os.EOL}${disambiguationText}`);
  }
  
  private async findExistingChatsByParticipants(expectedMemberEmails: string[]): Promise<Chat[]> {
    const chatType = expectedMemberEmails.length === 2 ? "oneOnOne" : "group";
    const endpoint = `${this.resource}/v1.0/chats?$filter=chatType eq '${chatType}'&$expand=members&$select=id,topic,createdDateTime,members`;
    const foundChats: Chat[] = [];

    const chats = await this.getAllChats(endpoint, []);

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

  private async findExistingGroupChatsByTopic(topic: string): Promise<Chat[]> {
    const endpoint = `${this.resource}/v1.0/chats?$filter=topic eq '${encodeURIComponent(topic)}'&$expand=members&$select=id,topic,createdDateTime,chatType`;
    const chats = await this.getAllChats(endpoint, []);
    return chats;
  }

  private async getAllChats(url: string, items: Chat[]): Promise<Chat[]> {
    const requestOptions: AxiosRequestConfig = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<ODataResponse<Chat>>(requestOptions);

    items = items.concat(res.value);

    if (res['@odata.nextLink']) {
      return await this.getAllChats(res['@odata.nextLink'] as string, items);
    }
    else {
      return items;
    }
  }
}

module.exports = new TeamsChatGetCommand();