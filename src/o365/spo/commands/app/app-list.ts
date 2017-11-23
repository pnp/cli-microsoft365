import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import {
  CommandHelp
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { AppMetadata } from './AppMetadata';
import Table = require('easy-table');

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.verbose)
      .then((accessToken: string): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Retrieved access token ${accessToken}. Loading apps from tenant app catalog...`);
        }

        cmd.log(`Retrieving apps...`);

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          }
        };

        if (this.verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: string): void => {
        if (this.verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const apps: { value: AppMetadata[] } = JSON.parse(res);

        if (apps.value && apps.value.length > 0) {
          const t: Table = new Table();
          apps.value.map((app: AppMetadata): void => {
            t.cell('Title', app.Title);
            t.cell('ID', app.ID);
            t.cell('Deployed', app.Deployed);
            t.cell('AppCatalogVersion', app.AppCatalogVersion);
            t.cell('InstalledVersion', app.InstalledVersion);
            t.newRow();
          });

          cmd.log('');
          cmd.log(t.toString());
        }
        else {
          cmd.log('No apps found');
        }
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public help(): CommandHelp {
    return function (args: {}, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.APP_LIST).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site, using the ${chalk.blue(commands.CONNECT)} command.
   
  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.APP_LIST}
      Returns the list of available apps from the tenant app catalog. Shows the installed version in the site if applicable.

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
    };
  }
}

module.exports = new AppListCommand();