import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  id: string;
  force?: boolean;
}

class TeamsTabRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TAB_REMOVE;
  }

  public get description(): string {
    return "Removes a tab from the specified channel";
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
        force: (!!args.options.force).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --teamId <teamId>'
      },
      {
        option: '-c, --channelId <channelId>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (!validation.isValidTeamsChannelId(args.options.channelId as string)) {
          return `${args.options.channelId} is not a valid Teams ChannelId`;
        }

        if (!validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTab = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}/channels/${args.options.channelId}/tabs/${formatting.encodeQueryParameter(args.options.id)}`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };
    if (args.options.force) {
      await removeTab();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the tab with id ${args.options.id} from channel ${args.options.channelId} in team ${args.options.teamId}?`
      });

      if (result.continue) {
        await removeTab();
      }
    }
  }
}

export default new TeamsTabRemoveCommand();