import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from './customaction';

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
    return commands.CUSTOMACTION_REMOVE;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
              logger.logToStderr(`Custom action with id ${args.options.id} not found`);
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeCustomAction();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${args.options.id} user custom action?`
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
      responseType: 'json'
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
        option: '-i, --id <id>'
      },
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
    if (validation.isValidGuid(args.options.id) === false) {
      return `${args.options.id} is not valid. Custom action Id (GUID) expected.`;
    }

    if (validation.isValidSharePointUrl(args.options.url) !== true) {
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
  }
}

module.exports = new SpoCustomActionRemoveCommand();