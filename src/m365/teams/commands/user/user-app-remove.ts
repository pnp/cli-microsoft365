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
  userId: string;
  confirm?: boolean;
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
        confirm: (!!args.options.confirm).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId <appId>'
      },
      {
        option: '--userId <userId>'
      },
      {
        option: '--confirm'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeApp: () => void = (): void => {
      const endpoint: string = `${this.resource}/v1.0`;

      const requestOptions: any = {
        url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps/${args.options.appId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .delete(requestOptions)
        .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
    };

    if (args.options.confirm) {
      removeApp();
    }
    else {
      Cli.prompt(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the app with id ${args.options.appId} for user ${args.options.userId}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeApp();
          }
        }
      );
    }
  }
}

module.exports = new TeamsUserAppRemoveCommand();