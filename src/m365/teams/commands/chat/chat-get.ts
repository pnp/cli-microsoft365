import { z } from 'zod';
import { Chat } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { chatUtil } from './chatUtil.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidTeamsChatId(val), {
    message: 'The value is not a valid Teams ChatId.'
  }).alias('i').optional(),
  participants: z.string().refine(val => {
    const participants = val.trim().toLowerCase().split(',').filter(e => e && e !== '');
    if (!participants || participants.length === 0) {
      return false;
    }
    return participants.every(e => validation.isValidUserPrincipalName(e));
  }, {
    message: 'The option participants contains one or more invalid email addresses.'
  }).alias('p').optional(),
  name: z.string().alias('n').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_GET;
  }

  public get description(): string {
    return 'Gets a Microsoft Teams chat conversation by id, participants or chat name.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.participants, opts.name].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of id, participants or name, but not more than one.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'participants', 'name']
        }
      });
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
    const currentUserEmail = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[this.resource].accessToken).toLowerCase();
    const existingChats = await chatUtil.findExistingChatsByParticipants([currentUserEmail, ...participants]);

    if (!existingChats || existingChats.length === 0) {
      throw 'No chat conversation was found with these participants.';
    }

    if (existingChats.length === 1) {
      return existingChats[0].id as string;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
    const result = await cli.handleMultipleResultsFound<Chat>(`Multiple chat conversations with these participants found.`, resultAsKeyValuePair);
    return result.id!;
  }

  private async getChatIdByName(name: string): Promise<string> {
    const existingChats = await chatUtil.findExistingGroupChatsByName(name);

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
}

export default new TeamsChatGetCommand();