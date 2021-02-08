import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  messageId: string;
}

class TeamsMessageGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAMS_MESSAGE_GET;
  }

  public get description(): string {
    return 'Retrieves a message from a channel in a Microsoft Teams team';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/beta/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
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
    if (!Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    return true;
  }
}

module.exports = new TeamsMessageGetCommand();