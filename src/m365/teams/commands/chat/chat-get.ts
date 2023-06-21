import { AadUserConversationMember, Chat, ConversationMember } from '@microsoft/microsoft-graph-types';
import os from 'os';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { chatUtil } from './chatUtil.js';
import { Cli } from '../../../../cli/Cli.js';

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
    return 'Get a Microsoft Teams chat conversation by id, participants or chat name.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        participants: typeof args.options.participants !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-p, --participants [participants]'
      },
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidTeamsChatId(args.options.id)) {
          return `${args.options.id} is not a valid Teams ChatId.`;
        }

        if (args.options.participants) {
          const participants = args.options.participants.trim().toLowerCase().split(',').filter(e => e && e !== '');
          if (!participants || participants.length === 0 || participants.some(e => !validation.isValidUserPrincipalName(e))) {
            return `${args.options.participants} contains one or more invalid email addresses.`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'participants', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const chatId = await this.getChatId(args);
      const chat: Chat = await this.getChatDetailsById(chatId as string);
      await logger.log(chat);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getChatId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    return args.options.participants
      ? this.getChatIdByParticipants(args.options.participants)
      : this.getChatIdByName(args.options.name as string);
  }

  private async getChatDetailsById(id: string): Promise<Chat> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/chats/${formatting.encodeQueryParameter(id)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<Chat>(requestOptions);
  }

  private async getChatIdByParticipants(participantsString: string): Promise<string> {
    const participants = participantsString.trim().toLowerCase().split(',').filter(e => e && e !== '');
    const currentUserEmail = accessToken.getUserNameFromAccessToken(auth.service.accessTokens[this.resource].accessToken).toLowerCase();
    const existingChats = await chatUtil.findExistingChatsByParticipants([currentUserEmail, ...participants]);

    if (!existingChats || existingChats.length === 0) {
      throw 'No chat conversation was found with these participants.';
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const disambiguationText = existingChats.map(c => {
      return `- ${c.id}${c.topic && ' - '}${c.topic} - ${c.createdDateTime && new Date(c.createdDateTime).toLocaleString()}`;
    }).join(os.EOL);

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
    return (await Cli.handleMultipleResultsFound<Chat>(`Multiple chat conversations with these participants found. Choose the correct ID:`, `Multiple chat conversations with these participants found. Please disambiguate:${os.EOL}${disambiguationText}`, resultAsKeyValuePair)).id!;
  }

  private async getChatIdByName(name: string): Promise<string> {
    const existingChats = await chatUtil.findExistingGroupChatsByName(name);

    if (!existingChats || existingChats.length === 0) {
      throw 'No chat conversation was found with this name.';
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const disambiguationText = existingChats.map(c => {
      const memberstring = (c.members as ConversationMember[]).map(m => (m as AadUserConversationMember).email).join(', ');
      return `- ${c.id} - ${c.createdDateTime && new Date(c.createdDateTime).toLocaleString()} - ${memberstring}`;
    }).join(os.EOL);

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
    return (await Cli.handleMultipleResultsFound<Chat>(`Multiple chat conversations with this name found. Choose the correct ID:`, `Multiple chat conversations with this name found. Please disambiguate:${os.EOL}${disambiguationText}`, resultAsKeyValuePair)).id!;
  }
}

export default new TeamsChatGetCommand();