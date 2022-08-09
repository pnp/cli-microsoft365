import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  teamId: string;
  confirm?: boolean;
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
        confirm: args.options.confirm || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId <appId>'
      },
      {
        option: '--teamId <teamId>'
      },
      {
        option: '--confirm'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const uninstallApp: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${args.options.teamId}/installedApps/${args.options.appId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      };

      request
        .delete(requestOptions)
        .then(_ => cb(), (res: Error): void => this.handleRejectedODataJsonPromise(res, logger, cb));
    };

    if (args.options.confirm) {
      uninstallApp();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to uninstall the app with id ${args.options.appId} from the Microsoft Teams team ${args.options.teamId}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          uninstallApp();
        }
      });
    }
  }
}

module.exports = new TeamsAppUninstallCommand();