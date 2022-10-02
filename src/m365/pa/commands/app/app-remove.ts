import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  confirm?: boolean;
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
        confirm: typeof args.options.confirm !== 'undefined'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--confirm'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing Microsoft Power App ${args.options.name}...`);
    }

    const removePaApp: () => Promise<void> = async (): Promise<void> => {
      const requestOptions: any = {
        url: `${this.resource}/providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2017-08-01`,
        resolveWithFullResponse: true,
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

    if (args.options.confirm) {
      await removePaApp();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?`
      });

      if (result.continue) {
        await removePaApp();
      }
    }
  }
}

module.exports = new PaAppRemoveCommand();