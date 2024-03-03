import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  force?: boolean;
  environmentName?: string;
  asAdmin?: boolean;
}

class PaAppRemoveCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Power App';
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
        force: typeof args.options.force !== 'undefined',
        asAdmin: !!args.options.asAdmin,
        environmentName: typeof args.options.environmentName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-f, --force'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-e, --environmentName [environmentName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }

        if (args.options.asAdmin && !args.options.environmentName) {
          return 'When specifying the asAdmin option, the environment option is required as well.';
        }

        if (args.options.environmentName && !args.options.asAdmin) {
          return 'When specifying the environment option, the asAdmin option is required as well.';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing Microsoft Power App ${args.options.name}...`);
    }

    const removePaApp = async (): Promise<void> => {
      let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
      if (args.options.asAdmin) {
        endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName!)}`;
      }
      endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2017-08-01`;

      const requestOptions: CliRequestOptions = {
        url: endpoint,
        fullResponse: true,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        if (err.response && err.response.status === 403) {
          throw new CommandError(`App '${args.options.name}' does not exist`);
        }
        else {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await removePaApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?` });

      if (result) {
        await removePaApp();
      }
    }
  }
}

export default new PaAppRemoveCommand();