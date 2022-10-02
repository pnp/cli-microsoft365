import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Message } from '../../Message';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  since?: string;
}

class TeamsMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_LIST;
  }

  public get description(): string {
    return 'Lists all messages from a channel in a Microsoft Teams team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'summary', 'body'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        since: typeof args.options.since !== 'undefined'
      });
    });
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
        option: '-s, --since [since]'
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

        if (args.options.since && !validation.isValidISODateDashOnly(args.options.since as string)) {
          return `${args.options.since} is not a valid ISO Date (with dash separator)`;
        }

        if (args.options.since && !validation.isDateInRange(args.options.since as string, 8)) {
          return `${args.options.since} is not in the last 8 months (for delta messages)`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deltaExtension: string = args.options.since !== undefined ? `/delta?$filter=lastModifiedDateTime gt ${args.options.since}` : '';
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages${deltaExtension}`;

    try {
      const items = await odata.getAllItems<Message>(endpoint);
      if (args.options.output !== 'json') {
        items.forEach(i => {
          i.body = i.body.content as any;
        });
      }

      logger.log(items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMessageListCommand();