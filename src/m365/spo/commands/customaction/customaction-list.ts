import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { CustomAction } from './customaction';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_LIST}`;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = args.options.scope ? args.options.scope : 'All';

    ((): Promise<CustomAction[]> => {
      if (this.debug) {
        cmd.log(`Attempt to get custom actions list with scope: ${scope}`);
        cmd.log('');
      }

      if (scope && scope.toLowerCase() !== "all") {
        return this.getCustomActions(args.options);
      }

      return this.searchAllScopes(args.options);
    })()
      .then((customActions: CustomAction[]): void => {
        if (customActions.length === 0) {
          if (this.verbose) {
            cmd.log(`Custom actions not found`);
          }
        }
        else {
          if (args.options.output === 'json') {
            cmd.log(customActions);
          }
          else {
            cmd.log(customActions.map(a => {
              return {
                Name: a.Name,
                Location: a.Location,
                Scope: this.humanizeScope(a.Scope),
                Id: a.Id
              };
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getCustomActions(options: Options): Promise<CustomAction[]> {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
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
        option: '-u, --url <url>',
        description: 'Url of the site (collection) to retrieve the custom action from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

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

module.exports = new SpoCustomActionListCommand();