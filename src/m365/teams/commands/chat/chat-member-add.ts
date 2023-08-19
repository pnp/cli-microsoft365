import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
  userId?: string;
  userName?: string;
  role?: string;
  visibleHistoryStartDateTime?: string;
  includeAllHistory?: boolean;
}

class TeamsChatMemberAddCommand extends GraphCommand {
  private static readonly roles: string[] = ['owner', 'guest'];

  public get name(): string {
    return commands.CHAT_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a member to a Microsoft Teams chat conversation.';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        role: typeof args.options.role !== 'undefined',
        visibleHistoryStartDateTime: typeof args.options.visibleHistoryStartDateTime !== 'undefined',
        includeAllHistory: !!args.options.includeAllHistory
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --chatId <chatId>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--role [role]',
        autocomplete: TeamsChatMemberAddCommand.roles
      },
      {
        option: '--visibleHistoryStartDateTime [visibleHistoryStartDateTime]'
      },
      {
        option: '--includeAllHistory'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
          return `${args.options.chatId} is not a valid chatId.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        if (args.options.role && TeamsChatMemberAddCommand.roles.indexOf(args.options.role) < 0) {
          return `${args.options.role} is not a valid role. Allowed values are ${TeamsChatMemberAddCommand.roles.join(', ')}`;
        }

        if (args.options.visibleHistoryStartDateTime && !validation.isValidISODateTime(args.options.visibleHistoryStartDateTime)) {
          return `'${args.options.visibleHistoryStartDateTime}' is not a valid visibleHistoryStartDateTime.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['userId', 'userName'] },
      {
        options: ['visibleHistoryStartDateTime', 'includeAllHistory'],
        runsWhen: (args) => args.options.visibleHistoryStartDateTime || args.options.includeAllHistory
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Adding member ${args.options.userId || args.options.userName} to chat with id ${args.options.chatId}...`);
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
          visibleHistoryStartDateTime: args.options.includeAllHistory ? '0001-01-01T00:00:00Z' : args.options.visibleHistoryStartDateTime,
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