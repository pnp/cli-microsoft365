import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  userId: string;
  force?: boolean;
}

class TeamsUserAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_REMOVE;
  }

  public get description(): string {
    return 'Uninstall an app from the personal scope of the specified user.';
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
        option: '--id <id>'
      },
      {
        option: '--userId <userId>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeApp = async (): Promise<void> => {
      const endpoint: string = `${this.resource}/v1.0`;

      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps/${args.options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
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
      await removeApp();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the app with id ${args.options.id} for user ${args.options.userId}?`
      });

      if (result.continue) {
        await removeApp();
      }
    }
  }
}

export default new TeamsUserAppRemoveCommand();