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
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  confirm?: boolean;
}

class SpoThemeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_REMOVE;
  }

  public get description(): string {
    return 'Removes existing theme';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeTheme = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}`);
          }

          if (this.verbose) {
            cmd.log(`Removing theme from tenant...`);
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_api/thememanager/DeleteTenantTheme`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${accessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            body: {
              name: args.options.name,
            },
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((rawRes: string): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(rawRes);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
    }

    if (args.options.confirm) {
      removeTheme();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the theme`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeTheme();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the theme to remove'
      },
      {
        option: '--confirm',
        description: 'Do not prompt for confirmation before removing theme'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To remove a theme, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
        
  Examples:

    Remove theme. Will prompt for confirmation before removing the theme
      ${chalk.grey(config.delimiter)} ${commands.THEME_REMOVE} --name Contoso-Blue
  
    Remove theme without prompting for confirmation
      ${chalk.grey(config.delimiter)} ${commands.THEME_REMOVE} --name Contoso-Blue --confirm

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview
      `);
  }
}

module.exports = new SpoThemeRemoveCommand();