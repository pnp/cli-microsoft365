import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from './customaction';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  scope?: string;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_LIST;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
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
        scope: args.options.scope || 'All'
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const scope: string = args.options.scope ? args.options.scope : 'All';

      if (this.debug) {
        logger.logToStderr(`Attempt to get custom actions list with scope: ${scope}`);
        logger.logToStderr('');
      }

      let customActions: CustomAction[];
      if (scope && scope.toLowerCase() !== "all") {
        customActions = await this.getCustomActions(args.options);
      }
      else {
        customActions = await this.searchAllScopes(args.options);
      }

      if (customActions.length === 0) {
        if (this.verbose) {
          logger.logToStderr(`Custom actions not found`);
        }
      }
      else {
        if (args.options.output !== 'json') {
          customActions.forEach(a => a.Scope = this.humanizeScope(a.Scope) as any);
        }

        logger.log(customActions);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private getCustomActions(options: Options): Promise<CustomAction[]> {
    const requestOptions: any = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return new Promise<CustomAction[]>((resolve: (list: CustomAction[]) => void, reject: (error: any) => void): void => {
      request
        .get<{ value: CustomAction[]; }>(requestOptions)
        .then((response: { value: CustomAction[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
  }

  /**
   * Two REST GET requests with `web` and `site` scope are sent.
   * The results are combined in one array.
   */
  private searchAllScopes(options: Options): Promise<CustomAction[]> {
    return new Promise<CustomAction[]>((resolve: (list: CustomAction[]) => void, reject: (error: any) => void): void => {
      options.scope = "Web";
      let webCustomActions: CustomAction[] = [];

      this
        .getCustomActions(options)
        .then((customActions: CustomAction[]): Promise<CustomAction[]> => {
          webCustomActions = customActions;

          options.scope = "Site";

          return this.getCustomActions(options);
        })
        .then((siteCustomActions: CustomAction[]): void => {
          resolve(siteCustomActions.concat(webCustomActions));
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }
}

module.exports = new SpoCustomActionListCommand();