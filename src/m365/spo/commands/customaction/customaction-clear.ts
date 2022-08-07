import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCustomActionClearCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_CLEAR;
  }

  public get description(): string {
    return 'Deletes all custom actions in the collection';
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
        scope: args.options.scope || 'All',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
          return isValidUrl;
        }
    
        if (args.options.scope &&
          args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const clearCustomActions = (): void => {
      ((): Promise<void> => {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.clearScopedCustomActions(args.options);
        }

        return this.clearAllScopes(args.options);
      })()
        .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      clearCustomActions();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear all the user custom actions with scope ${chalk.yellow(args.options.scope || 'All')}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          clearCustomActions();
        }
      });
    }
  }

  private clearScopedCustomActions(options: Options): Promise<void> {
    const requestOptions: any = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions/clear`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  /**
   * Clear request with `web` scope is send first. 
   * Another clear request is send with `site` scope after.
   */
  private clearAllScopes(options: Options): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .clearScopedCustomActions(options)
        .then((): Promise<void> => {
          options.scope = "Site";
          return this.clearScopedCustomActions(options);
        })
        .then((): void => {
          return resolve();
        }, (err: any): void => {
          reject(err);
        });
    });
  }
}

module.exports = new SpoCustomActionClearCommand();