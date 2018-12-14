import auth from '../../SpoAuth';
import { Auth } from '../../../../Auth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { SpoAppBaseCommand } from './app-base';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  appCatalogUrl?: string;
  skipFeatureDeployment?: boolean;
  scope?: string;
  siteUrl?: string;
}

class AppDeployCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_DEPLOY;
  }

  public get description(): string {
    return 'Deploys the specified app in the tenant app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.skipFeatureDeployment = args.options.skipFeatureDeployment || false;
    telemetryProps.scope = (!(!args.options.scope)).toString();
    telemetryProps.siteUrl = (!(!args.options.siteUrl)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let appId: string = '';
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let siteAccessToken: string = '';
    let appCatalogSiteUrl: string = '';

    this.getAppCatalogSiteUrl(cmd, auth.site.url, auth.service.accessToken, args)
      .then((siteUrl: string): Promise<string> => {
        appCatalogSiteUrl = siteUrl;

        const resource: string = Auth.getResourceFromUrl(appCatalogSiteUrl);
        return auth.getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug);
      })
      .then((accessToken: string): Promise<{ UniqueId: string }> | request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.verbose) {
          cmd.log('Retrieved access token');
        }

        if (args.options.id) {
          if (this.verbose) {
            cmd.log(`Using the specified app id ${args.options.id}`);
          }

          return Promise.resolve({ UniqueId: args.options.id });
        }
        else {
          if (this.verbose) {
            cmd.log(`Looking up app id for app named ${args.options.name}...`);
          }

          const requestOptions: any = {
            url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              accept: 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.get(requestOptions);
        }
      })
      .then((res: { UniqueId: string }): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        appId = res.UniqueId;

        if (this.verbose) {
          cmd.log(`Deploying app...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${appId}')/deploy`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata;charset=utf-8'
          }),
          body: { 'skipFeatureDeployment': args.options.skipFeatureDeployment || false },
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the app to deploy. Specify the id or the name but not both.'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the app to deploy. Specify the id or the name but not both.'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: '(optional) URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically'
      },
      {
        option: '--skipFeatureDeployment',
        description: 'If the app supports tenant-wide deployment, deploy it to the whole tenant'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Specify the target app catalog: \'tenant\' or \'sitecollection\' (default = tenant)',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--siteUrl [siteUrl]',
        description: 'The site url where the soultion package to be deployed. It must be specified when the scope is \'sitecollection\''
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      // verify either 'tenant' or 'site' specified if scope provided
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection' if specified`
        }

        if (testScope === 'sitecollection' && !args.options.siteUrl) {
          
          if(args.options.appCatalogUrl){
            return `You must specify siteUrl when the scope is sitecollection instead of appCatalogUrl`;
          }  
          return `You must specify siteUrl when the scope is sitecollection`;

        } else if(testScope === 'tenant' && args.options.siteUrl) {
          return `The siteUrl option can only be used when the scope option is set to sitecollection`;
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

      if (!args.options.scope && args.options.siteUrl) {
        return `The siteUrl option can only be used when the scope option is set to sitecollection`;
      }

      if(args.options.siteUrl) {
          return SpoAppBaseCommand.isValidSharePointUrl(args.options.siteUrl);
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_DEPLOY).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint site,
        using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To deploy an app in the tenant app catalog, you have to first log in to a SharePoint site
    using the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If you don't specify the URL of the tenant app catalog site using the ${chalk.grey('appCatalogUrl')} option,
    the CLI will try to determine its URL automatically. This will be done using SharePoint Search.
    If the tenant app catalog site hasn't been crawled yet, the CLI will not find it and will
    prompt you to provide the URL yourself.

    If the app with the specified ID doesn't exist in the tenant app catalog, the command will fail
    with an error. Before you can deploy an app, you have to add it to the tenant app catalog
    first using the ${chalk.blue(commands.APP_ADD)} command.
   
  Examples:
  
    Deploy the specified app in the tenant app catalog. Try to resolve the URL
    of the tenant app catalog automatically.
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --id 058140e3-0e37-44fc-a1d3-79c487d371a3

    Deploy the specified app in the site collection app catalog 
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}.
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --scope sitecollection --siteUrl https://contoso.sharepoint.com/sites/site1

    Deploy the app with the specified name in the tenant app catalog.
    Try to resolve the URL of the tenant app catalog automatically.
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --name solution.sppkg

    Deploy the specified app in the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --appCatalogUrl https://contoso.sharepoint.com/sites/apps

    Deploy the specified app to the whole tenant at once. Features included in the solution will not be activated.
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --skipFeatureDeployment
    
  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppDeployCommand();