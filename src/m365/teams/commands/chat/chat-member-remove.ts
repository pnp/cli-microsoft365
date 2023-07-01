import { ConversationMember } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Cli } from '../../../../cli/Cli.js';
import request, { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
  id?: string;
  userId?: string;
  userName?: string;
  force?: boolean;
}

class TeamsChatMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a member from a Microsoft Teams chat conversation';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --chatId <chatId>'
      },
      {
        option: '--id [id]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
          return `${args.options.chatId} is not a valid Teams chatId.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'userId', 'userName'] });
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
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}?`
      });

      if (result.continue) {
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