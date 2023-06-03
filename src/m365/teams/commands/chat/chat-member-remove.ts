import { ConversationMember } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import request, { CliRequestOptions } from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
  id?: string;
  userId?: string;
  userName?: string;
  confirm?: boolean;
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
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
    const removeUserFromChat: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.verbose) {
          logger.logToStderr(`Removing member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}...`);
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

    if (args.options.confirm) {
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

module.exports = new TeamsChatMemberRemoveCommand();