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
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  json: string;
}

class ThemeAddCommand extends SpoCommand {
  
  public get name(): string {
    return commands.THEME_ADD;
  }

  public get description(): string {
    return 'Adds new theme to the site with the given palette';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let siteAccessToken: string = '';
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest for tenant admin at ${auth.site.url}...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Adding theme to tenant...`);
        }

        let requestUrl: string = '';
        requestUrl = `${auth.site.url}/_api/thememanager/AddTenantTheme`;      

        const requestOptions: any = {
          url: requestUrl,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: {
            "name": args.options.name,
            "themeJson" : JSON.stringify(`{"palette":${args.options.json}}`),
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
      .then((themeResponse: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(themeResponse));
          cmd.log('');
        }

        cmd.log(JSON.stringify(themeResponse));
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));

  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--name <name>',
        description: 'name of the theme getting added'
      },
      {
        option: '--json <json>',
        description: 'color palette in the form of JSON object'
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

      if (!args.options.json) {
        return 'Required parameter json missing';
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
    
      To get information about a theme, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
          
    Example:
    
      Add new theme for the site
      ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
        ${chalk.grey(config.delimiter)} ${commands.THEME_ADD} --name Contoso-Blue --json <JSON object copied from URL>`);
  }

}

module.exports = new ThemeAddCommand();