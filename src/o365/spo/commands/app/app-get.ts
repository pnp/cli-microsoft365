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
import { AppMetadata } from './AppMetadata';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  appCatalogUrl?: string;
}

class AppGetCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the tenant app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.appCatalogUrl = typeof args.options.appCatalogUrl !== 'undefined';
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise | Promise<string> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}...`);
        }

        if (args.options.id) {
          return Promise.resolve(args.options.id);
        }
        else {
          let appCatalogUrl: string = '';

          return (args.options.appCatalogUrl ?
            Promise.resolve(args.options.appCatalogUrl) :
            this.getTenantAppCatalogUrl(cmd, this.debug))
            .then((appCatalogUrl: string): Promise<string> => {
              return Promise.resolve(appCatalogUrl);
            }, (error: any): Promise<string> => {
              if (this.debug) {
                cmd.log('Error');
                cmd.log(error);
                cmd.log('');
              }

              return new Promise<string>((resolve: (appCatalogUrl: string) => void, reject: (error: any) => void): void => {
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
            })
            .then((appCatalog: string): Promise<string> => {
              if (this.debug) {
                cmd.log(`Retrieved tenant app catalog URL ${appCatalog}`);
              }

              appCatalogUrl = appCatalog;

              if (this.verbose) {
                cmd.log(`Retrieving access token for the app catalog at ${appCatalogUrl}...`);
              }

              const appCatalogResource: string = Auth.getResourceFromUrl(appCatalogUrl);
              return auth.getAccessToken(appCatalogResource, auth.service.refreshToken as string, cmd, this.debug);
            })
            .then((token: string): request.RequestPromise => {
              if (this.verbose) {
                cmd.log(`Looking up app id for app named ${args.options.name}...`);
              }

              const requestOptions: any = {
                url: `${appCatalogUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
                headers: Utils.getRequestHeaders({
                  authorization: `Bearer ${token}`,
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
            })
            .then((res: { UniqueId: string }): Promise<string> => {
              if (this.debug) {
                cmd.log('Response:');
                cmd.log(res);
                cmd.log('');
              }

              return Promise.resolve(res.UniqueId);
            });
        }
      })
      .then((appId: string): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Retrieving information for app ${appId}...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps/GetById('${encodeURIComponent(appId)}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
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
      })
      .then((res: AppMetadata): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

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
        description: 'URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
        return SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
      using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To get information about the specified app available in the tenant app catalog,
    you have to first log in to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
   
  Examples:
  
    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the tenant app catalog.
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --id b2307a39-e878-458b-bc90-03bc578531d6

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog. Will try to detect the app catalog URL
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --name solution.sppkg

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog using the specified app catalog URL
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --name solution.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/apps

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppGetCommand();