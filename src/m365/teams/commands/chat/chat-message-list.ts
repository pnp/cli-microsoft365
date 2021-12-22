import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Message } from '../../Message';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  chatId: string;
}

class TeamsChatMessageListCommand extends GraphItemsListCommand<Message> {
  public get name(): string {
    return commands.CHAT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Lists all messages from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'summary', 'body'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;

    this
      .getAllItems(endpoint, logger, true)
      .then((): void => {
        if (args.options.output !== 'json') {
          this.items.forEach(i => {
            i.body = i.body.content as any;
          });
        }

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

module.exports = new TeamsChatMessageListCommand();