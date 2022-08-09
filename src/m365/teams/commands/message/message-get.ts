import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
    return commands.MESSAGE_GET;
  }

  public get description(): string {
    return 'Retrieves a message from a channel in a Microsoft Teams team';
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
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}`,
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
}

module.exports = new TeamsMessageGetCommand();