import { ItemBody, Message } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;

    try {
      const items = await odata.getAllItems<ExtendedMessage>(endpoint);
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
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChatMessageListCommand();