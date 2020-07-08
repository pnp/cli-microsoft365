import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { AppMetadata } from './AppMetadata';
import Utils from '../../../../Utils';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  id?: string;
  name?: string;
  scope?: string;
}

class SpoAppGetCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the specified app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.name = (!(!args.options.name)).toString();
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(cmd, spoUrl, args);
      })
      .then((appCatalogUrl: string): Promise<{ UniqueId: string }> => {
        appCatalogSiteUrl = appCatalogUrl;

        if (args.options.id) {
          return Promise.resolve({ UniqueId: args.options.id });
        }

        if (this.verbose) {
          cmd.log(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: { UniqueId: string }): Promise<AppMetadata> => {
        if (this.verbose) {
          cmd.log(`Retrieving information for app ${res}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(res.UniqueId)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: AppMetadata): void => {
        cmd.log(res);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant or site collection app catalog. It must be specified when the scope is \'sitecollection\''
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
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
          return `Scope must be either 'tenant' or 'sitecollection'`
        }

        if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
          return `You must specify appCatalogUrl when the scope is sitecollection`;
        }
      }

      if (!args.options.id && !args.options.name) {
        return 'Specify either the id or the name';
      }

      if (args.options.id && args.options.name) {
        return 'Specify either the id or the name but not both';
      }

      if (args.options.id && !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.appCatalogUrl) {
        return SpoAppBaseCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_GET).helpInformation());
    log(
      `  Remarks:
  
    When getting information about an app from the tenant app catalog,
    it's not necessary to specify the tenant app catalog URL. When the URL
    is not specified, the CLI will try to resolve the URL itself.
    Specifying the app catalog URL is required when you want to get information
    about an app from a site collection app catalog.

    When specifying site collection app catalog, you can specify the URL either
    with our without the ${chalk.grey('AppCatalog')} part, for example
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a/AppCatalog')} or
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a')}. CLI will accept both formats.
   
  Examples:
  
    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the tenant app catalog.
      ${commands.APP_GET} --id b2307a39-e878-458b-bc90-03bc578531d6

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog. Will try to detect the app catalog URL
      ${commands.APP_GET} --name solution.sppkg

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog using the specified app catalog URL
      ${commands.APP_GET} --name solution.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/apps

    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the site collection app catalog
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}.
      ${commands.APP_GET} --id b2307a39-e878-458b-bc90-03bc578531d6 --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppGetCommand();