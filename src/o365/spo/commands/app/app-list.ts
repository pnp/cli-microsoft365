import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import SpoCommand from '../../SpoCommand';
import { AppMetadata } from './AppMetadata';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class AppListCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the tenant app catalog';
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<string> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Loading apps from tenant app catalog...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving apps...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps`,
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

        return request.get(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const apps: { value: AppMetadata[] } = JSON.parse(res);

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
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site, using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:

    To list apps from the tenant app catalog, you have to first connect to a SharePoint site using
    the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    When using the text output type (default), the command lists only the values of the ${chalk.grey('Title')},
    ${chalk.grey('ID')}, ${chalk.grey('Deployed')} and ${chalk.grey('AppCatalogVersion')} properties of the app. When setting the output
    type to JSON, all available properties are included in the command output.
   
  Examples:
  
    Return the list of available apps from the tenant app catalog. Show the installed version in the site if applicable.
      ${chalk.grey(config.delimiter)} ${commands.APP_LIST}

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppListCommand();