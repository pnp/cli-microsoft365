import { cli } from '../../../../cli/cli.js';
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
  teamId: string;
  force?: boolean;
}

class TeamsAppUninstallCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_UNINSTALL;
  }

  public get description(): string {
    return 'Uninstalls an app from a Microsoft Team team';
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
        force: args.options.force || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--teamId <teamId>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const uninstallApp = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Uninstalling app with ID ${args.options.id} in team ${args.options.teamId}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${args.options.teamId}/installedApps/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await uninstallApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to uninstall the app with id ${args.options.id} from the Microsoft Teams team ${args.options.teamId}?` });

      if (result) {
        await uninstallApp();
      }
    }
  }
}

export default new TeamsAppUninstallCommand();