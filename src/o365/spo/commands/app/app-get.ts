import auth from '../../SpoAuth';
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
  id: string;
}

class AppGetCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the tenant app catalog';
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
          cmd.log(`Retrieving app information...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')`,
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

        const app: AppMetadata = JSON.parse(res);
        cmd.log(app);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-i, --id <id>',
      description: 'ID of the app to retrieve information for'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To get information about the specified app available in the tenant app catalog,
    you have to first connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
   
  Examples:
  
    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the tenant app catalog.
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} -i b2307a39-e878-458b-bc90-03bc578531d6

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppGetCommand();