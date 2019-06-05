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

const vorpal: Vorpal = require('../../../../vorpal-init');

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

      if (!args.options.id) {
        return 'Required parameter id missing';
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

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_REMOVE).helpInformation());
    log(
      `  Remarks:
  
    When removing an app from the tenant app catalog, it's not necessary
    to specify the tenant app catalog URL. When the URL is not specified,
    the CLI will try to resolve the URL itself. Specifying the app catalog URL
    is required when you want to remove the app from a site collection
    app catalog.

    When specifying site collection app catalog, you can specify the URL either
    with our without the ${chalk.grey('AppCatalog')} part, for example
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a/AppCatalog')} or
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a')}. CLI will accept both formats.

    If the app with the specified ID doesn't exist in the tenant app catalog,
    the command will fail with an error.
   
  Examples:
  
    Remove the specified app from the tenant app catalog. Try to resolve the URL
    of the tenant app catalog automatically. Additionally, will prompt for
    confirmation before actually removing the app.
      ${commands.APP_REMOVE} --id 058140e3-0e37-44fc-a1d3-79c487d371a3

    Remove the specified app from the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}. Additionally, will prompt
    for confirmation before actually retracting the app.
      ${commands.APP_REMOVE} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --appCatalogUrl https://contoso.sharepoint.com/sites/apps

    Remove the specified app from the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}. Don't prompt for confirmation.
      ${commands.APP_REMOVE} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --appCatalogUrl https://contoso.sharepoint.com/sites/apps --confirm

    Remove the specified app from a site collection app catalog 
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}.
      ${commands.APP_REMOVE} --id d95f8c94-67a1-4615-9af8-361ad33be93c --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1
    
  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppRemoveCommand();