import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CustomAction } from './customaction';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  url: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCustomActionRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified custom action';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeCustomAction = (): void => {
      ((): Promise<CustomAction | void> => {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.removeScopedCustomAction(args.options);
        }

        return this.searchAllScopes(args.options);
      })()
        .then((customAction: CustomAction | void): void => {
          if (this.verbose) {
            if (customAction && customAction["odata.null"] === true) {
              cmd.log(`Custom action with id ${args.options.id} not found`);
            }
            else {
              cmd.log(chalk.green('DONE'));
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      removeCustomAction();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${args.options.id} user custom action?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeCustomAction();
        }
      });
    }
  }

  private removeScopedCustomAction(options: Options): Promise<undefined> {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions('${encodeURIComponent(options.id)}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE'
      },
      json: true
    };

    return request.post(requestOptions);
  }

  /**
   * Remove request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private searchAllScopes(options: Options): Promise<CustomAction | undefined> {
    return new Promise<CustomAction | undefined>((resolve: (result: CustomAction | undefined) => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .removeScopedCustomAction(options)
        .then((webResult: CustomAction | undefined): void => {
          if (webResult === undefined) {
            return resolve(webResult);
          }

          options.scope = "Site";
          this
            .removeScopedCustomAction(options)
            .then((siteResult: CustomAction | undefined): void => {
              return resolve(siteResult);
            }, (err: any): void => {
              reject(err);
            });
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'Id (GUID) of the custom action to remove'
      },
      {
        option: '-u, --url <url>',
        description: 'Url of the site or site collection to remove the custom action from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of a user custom action'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (Utils.isValidGuid(args.options.id) === false) {
        return `${args.options.id} is not valid. Custom action Id (GUID) expected.`;
      }

      if (SpoCommand.isValidSharePointUrl(args.options.url) !== true) {
        return 'Missing required option url';
      }

      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoCustomActionRemoveCommand();