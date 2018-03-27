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
import { ContextInfo } from '../../spo';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  appCatalogUrl?: string;
  skipFeatureDeployment?: boolean;
}

class AppDeployCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_DEPLOY;
  }

  public get description(): string {
    return 'Deploys the specified app in the tenant app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.skipFeatureDeployment = args.options.skipFeatureDeployment || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let appCatalogUrl: string = '';
    let accessToken: string = '';

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<string> => {
        return new Promise<string>((resolve: (appCatalogUrl: string) => void, reject: (error: any) => void): void => {
          if (args.options.appCatalogUrl) {
            resolve(args.options.appCatalogUrl);
          }
          else {
            this
              .getTenantAppCatalogUrl(cmd, this.debug)
              .then((appCatalogUrl: string): void => {
                resolve(appCatalogUrl);
              }, (error: any): void => {
                if (this.debug) {
                  cmd.log('Error');
                  cmd.log(error);
                  cmd.log('');
                }

                cmd.log('CLI could not automatically determine the URL of the tenant app catalog');
                cmd.log('What is the absolute URL of your tenant app catalog site');
                cmd.prompt({
                  type: 'input',
                  name: 'appCatalogUrl',
                  message: '? ',
                }, (result: { appCatalogUrl?: string }): void => {
                  if (!result.appCatalogUrl) {
                    reject(`Couldn't determine tenant app catalog URL`);
                  }
                  else {
                    let isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(result.appCatalogUrl);
                    if (isValidSharePointUrl === true) {
                      resolve(result.appCatalogUrl);
                    }
                    else {
                      reject(isValidSharePointUrl);
                    }
                  }
                });
              });
          }
        });
      })
      .then((appCatalog: string): Promise<string> => {
        if (this.debug) {
          cmd.log(`Retrieved tenant app catalog URL ${appCatalog}`);
        }

        appCatalogUrl = appCatalog;

        let appCatalogResource: string = Auth.getResourceFromUrl(appCatalog);
        return auth.getAccessToken(appCatalogResource, auth.service.refreshToken as string, cmd, this.debug);
      })
      .then((token: string): request.RequestPromise => {
        accessToken = token;

        return this.getRequestDigestForSite(appCatalogUrl, accessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Deploying app...`);
        }

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/tenantappcatalog/AvailableApps/GetById('${args.options.id}')/deploy`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata;charset=utf-8',
            'X-RequestDigest': res.FormDigestValue
          }),
          body: JSON.stringify({ 'skipFeatureDeployment': args.options.skipFeatureDeployment || false })
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
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to deploy. Needs to be available in the tenant app catalog.'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: '(optional) URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically'
      },
      {
        option: '--skipFeatureDeployment',
        description: 'If the app supports tenant-wide deployment, deploy it to the whole tenant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (args.options.appCatalogUrl) {
        return SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_DEPLOY).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint site,
        using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To deploy an app in the tenant app catalog, you have to first connect to a SharePoint site
    using the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

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
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} -i 058140e3-0e37-44fc-a1d3-79c487d371a3

    Deploy the specified app in the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/apps

    Deploy the specified app to the whole tenant at once. Features included in the solution will not be activated.
      ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 --skipFeatureDeployment
    
  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppDeployCommand();