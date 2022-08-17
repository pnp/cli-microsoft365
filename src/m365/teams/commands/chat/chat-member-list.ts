import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;

    odata
      .getAllItems(endpoint)
      .then((items): void => {
        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsChatMemberListCommand();