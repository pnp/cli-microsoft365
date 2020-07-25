import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';
import SpoCommand from '../../../base/SpoCommand';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  confirm?: boolean;
  id: string;
  scope?: string;
}

class SpoAppRemoveCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified app from the specified app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    const removeApp: () => void = (): void => {
      this
        .getSpoUrl(cmd, this.debug)
        .then((spoUrl: string): Promise<string> => {
          return this.getAppCatalogSiteUrl(cmd, spoUrl, args)
        })
        .then((appCatalogUrl: string): Promise<void> => {
          if (this.debug) {
            cmd.log(`Retrieved app catalog URL ${appCatalogUrl}. Removing app from the app catalog...`);
          }

          const requestOptions: any = {
            url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/remove`,
            headers: {
              accept: 'application/json;odata=nometadata'
            }
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          cb();
        }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
    };

    if (args.options.confirm) {
      removeApp();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the app ${args.options.id} from the app catalog?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to remove. Needs to be available in the tenant app catalog.'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant or site collection app catalog. It must be specified when the scope is \'sitecollection\''
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      // verify either 'tenant' or 'sitecollection' specified if scope provided
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection' if specified`
        }

        if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
          return `You must specify appCatalogUrl when the scope is sitecollection`;
        }
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.appCatalogUrl) {
        return SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      return true;
    };
  }
}

module.exports = new SpoAppRemoveCommand();