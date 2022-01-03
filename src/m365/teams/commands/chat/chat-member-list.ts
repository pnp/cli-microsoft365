import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
}

class TeamsChatMemberListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.CHAT_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists all members from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['userId', 'displayName', 'email'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;

    this
      .getAllItems(endpoint, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --chatId <chatId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidTeamsChatId(args.options.chatId)) {
      return `${args.options.chatId} is not a valid Teams ChatId`;
    }

    return true;
  }
}

module.exports = new TeamsChatMemberListCommand();