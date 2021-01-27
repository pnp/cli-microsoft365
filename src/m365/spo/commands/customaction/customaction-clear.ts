import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCustomActionClearCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_CLEAR}`;
  }

  public get description(): string {
    return 'Deletes all custom actions in the collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const clearCustomActions = (): void => {
      ((): Promise<void> => {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.clearScopedCustomActions(args.options);
        }

        return this.clearAllScopes(args.options);
      })()
        .then((): void => {
          if (this.verbose) {
            logger.logToStderr(chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    }

    if (args.options.confirm) {
      clearCustomActions();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear all the user custom actions with scope ${chalk.yellow(args.options.scope || 'All')}?`,
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
      url: `${options.url}/_api/${options.scope}/UserCustomActions/clear`,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.url) {
      return 'Missing required option url';
    }

    const isValidUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
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
}

module.exports = new SpoCustomActionClearCommand();