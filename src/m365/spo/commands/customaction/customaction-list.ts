import { Logger } from '../../../../cli';
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
  url: string;
  scope?: string;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_LIST;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = args.options.scope ? args.options.scope : 'All';

    ((): Promise<CustomAction[]> => {
      if (this.debug) {
        logger.logToStderr(`Attempt to get custom actions list with scope: ${scope}`);
        logger.logToStderr('');
      }

      if (scope && scope.toLowerCase() !== "all") {
        return this.getCustomActions(args.options);
      }

      return this.searchAllScopes(args.options);
    })()
      .then((customActions: CustomAction[]): void => {
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
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getCustomActions(options: Options): Promise<CustomAction[]> {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions`,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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

module.exports = new SpoCustomActionListCommand();