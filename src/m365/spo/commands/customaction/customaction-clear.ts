import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clearCustomActions = (): void => {
      ((): Promise<void> => {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.clearScopedCustomActions(args.options);
        }

        return this.clearAllScopes(args.options);
      })()
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      clearCustomActions();
    }
    else {
      cmd.prompt({
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
      json: true
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
        option: '-u, --url <url>',
        description: 'Url of the site or site collection to clear the custom actions from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing all custom actions'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }
}

module.exports = new SpoCustomActionClearCommand();