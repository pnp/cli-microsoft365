import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
}

class TeamsChatMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists all members from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['userId', 'displayName', 'email'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --chatId <chatId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
          return `${args.options.chatId} is not a valid Teams ChatId`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;

    try {
      const items = await odata.getAllItems(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatMemberListCommand();