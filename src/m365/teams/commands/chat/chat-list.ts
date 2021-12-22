import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
}

class TeamsChatListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.CHAT_LIST;
  }

  public get description(): string {
    return 'Lists all chat conversations';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'topic', 'chatType'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const filter = args.options.type !== undefined ? `?$filter=chatType eq '${args.options.type}'` : '';
    const endpoint: string = `${this.resource}/v1.0/chats${filter}`;

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
        option: '-t, --type [chatType]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }


  public validate(args: CommandArgs): boolean | string {
    const supportedTypes = ['oneOnOne', 'group', 'meeting'];
    if (args.options.type !== undefined && supportedTypes.indexOf(args.options.type) === -1) {
      return `${args.options.type} is not a valid chatType. Accepted values are ${supportedTypes.join(', ')}`;
    }

    return true;
  }
}

module.exports = new TeamsChatListCommand();