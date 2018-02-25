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
}

class ThemeRemoveCommand extends SpoCommand {

  public get name(): string {
    return commands.THEME_REMOVE;
  }

  public get description(): string {
    return 'Removes existing theme from tenant with the given name.';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Removing theme from tenant store...`);
        }

        if (this.verbose) {
          cmd.log(`Removing theme from tenant...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/thememanager/DeleteTenantTheme`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: {
            "name": args.options.name,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
        option: '-n, --name <name>',
        description: 'name of the theme getting removed'
    }];

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
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.
  
    Remarks:
    
      To remove the theme, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
          
    Example:
    
      Removes theme from tenant
      ${chalk.grey(config.delimiter)} ${commands.THEME_REMOVE} -n Contoso-Blue`);
  }
}

module.exports = new ThemeRemoveCommand();