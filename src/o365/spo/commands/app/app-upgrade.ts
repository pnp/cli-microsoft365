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

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  identity: string;
}

class AppUpgradeCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades an installed app';
  }

  
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.AppId = args.options.identity;
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
          cmd.log(`Retrieved access token ${accessToken}.`);
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

        cmd.log(`Installing app...`);

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/AvailableApps/GetByID('${args.options.identity}')/upgrade`,
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

        return request.post(requestOptions);
      })
      .then((res: string): void => {
     
        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log("App succesfully installed");
        
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-id, --identity <guid>',
        description: 'The id of the app to deploy. Needs to be available in the app catalog.'
      }
    ];

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
      if (!args.options.identity) {
        return `identity is a required parameter.`;
      }
      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.APP_UPGRADE).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
  using the ${chalk.blue(commands.CONNECT)} command.
   
  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.APP_UPGRADE} --identity 058140e3-0e37-44fc-a1d3-79c487d371a3
      Upgrades a given up if a newer version is available.

`);
    };
  }
}

module.exports = new AppUpgradeCommand();