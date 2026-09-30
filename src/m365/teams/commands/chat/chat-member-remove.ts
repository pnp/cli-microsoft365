import { z } from 'zod';
import { ConversationMember } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string().refine(val => validation.isValidTeamsChatId(val), {
    message: 'The value of the option chatId must be a valid Teams ChatId.'
  }).alias('i'),
  id: z.string().optional(),
  userId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value of the option userId must be a valid GUID.'
  }).optional(),
  userName: z.string().refine(val => validation.isValidUserPrincipalName(val), {
    message: 'The value of the option userName must be a valid user principal name.'
  }).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChatMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a member from a Microsoft Teams chat conversation';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.userId, opts.userName].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of id, userId or userName, but not more than one.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeUserFromChat = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}...`);
        }

        const memberId = await this.getMemberId(args);
        const chatMemberRemoveOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/chats/${args.options.chatId}/members/${memberId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };
        await request.delete(chatMemberRemoveOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeUserFromChat();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}?` });

      if (result) {
        await removeUserFromChat();
      }
    }
  }

  private async getMemberId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const memberRequestUrl: string = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;
    const members = await odata.getAllItems<ConversationMember>(memberRequestUrl);
    if (args.options.userName) {
      const matchingMember: any = members.find((memb: any) => memb.email.toLowerCase() === args.options.userName!.toLowerCase());
      if (!matchingMember) {
        throw `Member with userName '${args.options.userName}' could not be found in the chat.`;
      }
      return matchingMember.id;
    }
    else {
      const matchingMember: any = members.find((memb: any) => memb.userId.toLowerCase() === args.options.userId!.toLowerCase());
      if (!matchingMember) {
        throw `Member with userId '${args.options.userId}' could not be found in the chat.`;
      }
      return matchingMember.id;
    }
  }
}

export default new TeamsChatMemberRemoveCommand();