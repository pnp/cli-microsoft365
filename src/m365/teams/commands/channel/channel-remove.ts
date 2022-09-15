import { Cli, Logger } from '../../../../cli';
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
  channelId?: string;
  channelName?: string;
  teamId: string;
  confirm?: boolean;
}

class TeamsChannelRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified channel in the Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --channelId [channelId]'
      },
      {
        option: '-n, --channelName [channelName]'
      },
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
          return `${args.options.channelId} is not a valid Teams Channel Id`;
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['channelId', 'channelName']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeChannel: () => Promise<void> = async (): Promise<void> => {
      try {
        if (args.options.channelName) {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName)}'`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };
  
          const res: { value: Channel[] } = await request.get<{ value: Channel[] }>(requestOptions);
          const channelItem: Channel | undefined = res.value[0];

          if (!channelItem) {
            return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
          }

          const channelId: string = res.value[0].id;

          const requestOptionsDelete: any = {
            url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(channelId)}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          await request.delete(requestOptionsDelete);
        }
  
        if (args.options.channelId) {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };
  
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeChannel();
    }
    else {
      const channelName = args.options.channelName ? args.options.channelName : args.options.channelId;
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the channel ${channelName} from team ${args.options.teamId}?`
      });
      
      if (result.continue) {
        await removeChannel();
      }
    }
  }
}

module.exports = new TeamsChannelRemoveCommand();