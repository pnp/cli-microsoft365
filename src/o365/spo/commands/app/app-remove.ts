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
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  appCatalogUrl?: string;
  confirm?: boolean;
}

class AppRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified app from the tenant app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let appCatalogUrl: string = '';

    const removeApp: () => void = (): void => {
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
        .then((accessToken: string): request.RequestPromise => {
          if (this.debug) {
            cmd.log(`Retrieved access token for the tenant app catalog ${accessToken}. Removing app from the app catalog...`);
          }

          const requestOptions: any = {
            url: `${appCatalogUrl}/_api/web/tenantappcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/remove`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${accessToken}`,
              accept: 'application/json;odata=nometadata'
            })
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

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
        message: `Are you sure you want to remove the app ${args.options.id} from the tenant app catalog?`,
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
        description: '(optional) URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically'
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
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.appCatalogUrl) {
        const isValidSharePointUrl: string | boolean = SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_REMOVE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint site,
        using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To remove an app from the tenant app catalog, you have to first connect to a SharePoint site
    using the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If you don't specify the URL of the tenant app catalog site using the ${chalk.grey('appCatalogUrl')} option,
    the CLI will try to determine its URL automatically. This will be done using SharePoint Search.
    If the tenant app catalog site hasn't been crawled yet, the CLI will not find it and will
    prompt you to provide the URL yourself.

    If the app with the specified ID doesn't exist in the tenant app catalog, the command will fail
    with an error.
   
  Examples:
  
    Remove the specified app from the tenant app catalog. Try to resolve the URL
    of the tenant app catalog automatically. Additionally, will prompt for confirmation before
    actually removing the app.
      ${chalk.grey(config.delimiter)} ${commands.APP_REMOVE} --id 058140e3-0e37-44fc-a1d3-79c487d371a3

    Remove the specified app from the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}. Additionally, will prompt for confirmation before
    actually retracting the app.
      ${chalk.grey(config.delimiter)} ${commands.APP_REMOVE} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/apps

    Remove the specified app from the tenant app catalog located at
    ${chalk.grey('https://contoso.sharepoint.com/sites/apps')}. Don't prompt for confirmation.
      ${chalk.grey(config.delimiter)} ${commands.APP_REMOVE} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/apps --confirm
    
  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppRemoveCommand();