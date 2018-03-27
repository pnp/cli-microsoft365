import { ContextInfo } from './../../spo';
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
  siteUrl: string;
  confirm?: boolean;
}

class AppUninstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UNINSTALL;
  }

  public get description(): string {
    return 'Uninstalls an app from the site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const uninstallApp: () => void = (): void => {
      const resource: string = Auth.getResourceFromUrl(args.options.siteUrl);
      let siteAccessToken: string = '';

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          siteAccessToken = accessToken;

          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
          }

          return this.getRequestDigestForSite(args.options.siteUrl, siteAccessToken, cmd, this.debug);
        })
        .then((res: ContextInfo): request.RequestPromise => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
          }

          const requestOptions: any = {
            url: `${args.options.siteUrl}/_api/web/tenantappcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/uninstall`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              accept: 'application/json;odata=nometadata',
              'X-RequestDigest': res.FormDigestValue
            })
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

          cb();
        }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
    };

    if (args.options.confirm) {
      uninstallApp();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to uninstall the app ${args.options.id} from site ${args.options.siteUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          uninstallApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to retrieve information for'
      },
      {
        option: '-s, --siteUrl <siteUrl>',
        description: 'Absolute URL of the site to install the app in'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming uninstalling the app'
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

      if (!args.options.siteUrl) {
        return 'Required parameter siteUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.siteUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_UNINSTALL).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To uninstall an app from the site, you have to first connect to a SharePoint site using
    the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
   
  Examples:
  
    Uninstall the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    from the ${chalk.grey('https://contoso.sharepoint.com')} site.
      ${chalk.grey(config.delimiter)} ${commands.APP_UNINSTALL} -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com

    Uninstall the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    from the ${chalk.grey('https://contoso.sharepoint.com')} site without prompting for confirmation.
      ${chalk.grey(config.delimiter)} ${commands.APP_UNINSTALL} -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com --confirm

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppUninstallCommand();