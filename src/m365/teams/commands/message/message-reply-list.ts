import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Reply } from '../../Reply';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  messageId: string;
}

class TeamsMessageReplyListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_REPLY_LIST;
  }

  public get description(): string {
    return 'Retrieves replies to a message from a channel in a Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'body'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '-c, --channelId <channelId>'
      },
      {
        option: '-m, --messageId <messageId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (!validation.isValidTeamsChannelId(args.options.channelId as string)) {
          return `${args.options.channelId} is not a valid Teams ChannelId`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}/replies`;

    odata
      .getAllItems<Reply>(endpoint)
      .then((items): void => {
        if (args.options.output !== 'json') {
          items.forEach(i => {
            i.body = i.body.content as any;
          });
        }

        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsMessageReplyListCommand();