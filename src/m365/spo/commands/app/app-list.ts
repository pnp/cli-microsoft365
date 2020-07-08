import commands from '../../commands';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { AppMetadata } from './AppMetadata';
import GlobalOptions from '../../../../GlobalOptions';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  scope?: string;
  appCatalogUrl?: string;
}

class SpoAppListCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the specified app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<string> => {
        spoUrl = _spoUrl;
        return this.getAppCatalogSiteUrl(cmd, spoUrl, args)
      })
      .then((appCatalogUrl: string): Promise<{ value: AppMetadata[] }> => {
        appCatalogSiteUrl = appCatalogUrl;

        if (this.verbose) {
          cmd.log(`Retrieving apps...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((apps: { value: AppMetadata[] }): void => {
        if (apps.value && apps.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(apps.value);
          }
          else {
            cmd.log(apps.value.map(a => {
              return {
                Title: a.Title,
                ID: a.ID,
                Deployed: a.Deployed,
                AppCatalogVersion: a.AppCatalogVersion
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No apps found');
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant or site collection app catalog. It must be specified when the scope is \'sitecollection\''
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
          return `Scope must be either 'tenant' or 'sitecollection'`;
        }

        if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
          return `You must specify appCatalogUrl when the scope is sitecollection`;
        }

        if (args.options.appCatalogUrl) {
          return SpoAppBaseCommand.isValidSharePointUrl(args.options.appCatalogUrl);
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_LIST).helpInformation());
    log(
      `  Remarks:

    When listing information about apps available in the tenant app catalog,
    it's not necessary to specify the tenant app catalog URL. When the URL
    is not specified, the CLI will try to resolve the URL itself.
    Specifying the app catalog URL is required when you want to list information
    about apps in a site collection app catalog.

    When specifying site collection app catalog, you can specify the URL either
    with our without the ${chalk.grey('AppCatalog')} part, for example
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a/AppCatalog')} or
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a')}. CLI will accept both formats.

    When using the text output type (default), the command lists only the values
    of the ${chalk.grey('Title')}, ${chalk.grey('ID')}, ${chalk.grey('Deployed')} and ${chalk.grey('AppCatalogVersion')} properties of the app.
    When setting the output type to JSON, all available properties are included
    in the command output.
   
  Examples:
  
    Return the list of available apps from the tenant app catalog.
    Show the installed version in the site if applicable.
      ${commands.APP_LIST}

    Return the list of available apps from a site collection app catalog
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}.
      ${commands.APP_LIST} --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppListCommand();