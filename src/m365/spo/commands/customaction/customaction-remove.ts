import { Cli, Logger } from '../../../../cli';
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
  id?: string;
  title?: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
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
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && validation.isValidGuid(args.options.id) === false) {
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
    );
  }

  #initOptionSets(): void {
  	this.optionSets.push(['id', 'title']);
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

  private getCustomActionId(options: Options): Promise<string> {
    if (options.id) {
      return Promise.resolve(options.id);
    }

    const customActionRequestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions?$filter=Title eq '${encodeURIComponent(options.title as string)}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: CustomAction[] }>(customActionRequestOptions)
      .then((res: { value: CustomAction[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].Id);
        }

        if (res.value.length === 0) {
          return Promise.reject(`No user custom action with title '${options.title}' found`);
        }

        return Promise.reject(`Multiple user custom actions with title '${options.title}' found. Please disambiguate using IDs: ${res.value.map(a => a.Id).join(', ')}`);
      });
  }

  private removeScopedCustomAction(options: Options): Promise<undefined> {
    return this
      .getCustomActionId(options)
      .then((customActionId: string): Promise<undefined> => {
        const requestOptions: any = {
          url: `${options.url}/_api/${options.scope}/UserCustomActions('${encodeURIComponent(customActionId)}')')`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'X-HTTP-Method': 'DELETE'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      });
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
}

module.exports = new SpoCustomActionRemoveCommand();