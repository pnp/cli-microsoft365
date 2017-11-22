import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import * as request from 'request-promise-native';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { ContextInfo } from '../../spo';
import SpoCommand from '../../SpoCommand';
import { RestResponse } from '../../models/RestResponse';
import { AppMetadata } from '../../models/AppMetadata';

const vorpal: Vorpal = require('../../../../vorpal-init');
const Table = require('easy-table');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
}

class AppListCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Retrieves the apps from the tenant app catalog';
  }

  
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const verbose: boolean = args.options.verbose || false;


    auth
      .ensureAccessToken(auth.service.resource, this, verbose)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (verbose) {
          cmd.log(`Retrieved access token ${accessToken}. Loading CDN settings for the ${auth.site.url} tenant...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/contextinfo`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Retrieving apps...`);

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=verbose',
            'X-RequestDigest': res.FormDigestValue
          },
          body: ``
        };

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: string): void => {
     
        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: RestResponse<AppMetadata> = JSON.parse(res);

        const apps = json.d.results;      

        var t: any = new Table();
        
        for (let app of apps) {
          t.cell('Title',app.Title);
          t.cell('ID',app.ID);
          t.cell('Deployed',app.Deployed);
          t.cell('AppCatalogVersion',app.AppCatalogVersion);
          t.cell('InstalledVersion',app.InstalledVersion);
          t.newRow();
        }
        cmd.log(t.toString());
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [];

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.APP_LIST).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
   
  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.APP_LIST}
      Returns the list of available apps from the tenant app catalog. Shows the installed version in the site if applicable.

`);
    };
  }
}

module.exports = new AppListCommand();