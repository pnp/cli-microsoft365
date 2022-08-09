import { ItemBody, Message } from '@microsoft/microsoft-graph-types';
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

interface ExtendedMessage extends Message {
  shortBody?: string;
}

class TeamsChatMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Lists all messages from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'shortBody'];
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
          return `${args.options.chatId} is not a valid Teams chat ID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;

    odata
      .getAllItems<ExtendedMessage>(endpoint)
      .then((items): void => {
        if (args.options.output !== 'json') {
          items.forEach(i => {
            // hoist the content to body for readability
            i.body = (i.body as ItemBody).content as any;

            let shortBody: string | undefined;
            const bodyToProcess = i.body as string;

            if (bodyToProcess) {
              let maxLength = 50;
              let addedDots = '...';
              if (bodyToProcess.length < maxLength) {
                maxLength = bodyToProcess.length;
                addedDots = '';
              }

              shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
            }

            i.shortBody = shortBody;
          });
        }

        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsChatMessageListCommand();