import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelName: string;
  description?: string
  newChannelName?: string;
  teamId: string;
}

class TeamsChannelSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_SET;
  }
  public get description(): string {
    return 'Updates properties of the specified channel in the given Microsoft Teams team';
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
        newChannelName: typeof args.options.newChannelName !== 'undefined',
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--channelName <channelName>'
      },
      {
        option: '--newChannelName [newChannelName]'
      },
      {
        option: '--description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelName.toLowerCase() === "general") {
          return 'General channel cannot be updated';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ value: Channel[] }>(requestOptions)
      .then((res: { value: Channel[] }): Promise<void> => {
        const channelItem: Channel | undefined = res.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        const channelId: string = res.value[0].id;
        const data: any = this.mapRequestBody(args.options);
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${channelId}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: data
        };

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.newChannelName) {
      requestBody.displayName = options.newChannelName;
    }

    if (options.description) {
      requestBody.description = options.description;
    }

    return requestBody;
  }
}

module.exports = new TeamsChannelSetCommand();