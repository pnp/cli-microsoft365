import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

class TeamsMessageReplyListCommand extends GraphCommand  {
  public get name(): string {
    return commands.MESSAGE_REPLY_LIST;
  }

  public get description(): string {
    return 'Retrieves replies to a message from a channel in a Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'body'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}/replies`;

    odata
      .getAllItems<Reply>(endpoint, logger)
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '-c, --channelId <channelId>'
      },
      {
        option: '-m, --messageId <messageId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (!validation.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    return true;
  }
}

module.exports = new TeamsMessageReplyListCommand();