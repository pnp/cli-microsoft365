import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class SpoThemeListCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_LIST;
  }

  public get description(): string {
    return 'Retrieves the list of custom themes';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving themes from tenant store...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/thememanager/GetTenantThemingOptions`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((rawRes: any): void => {
        if (args.options.debug) {
          cmd.log('Response:');
          cmd.log(rawRes);
          cmd.log('');
        }

        const themePreviews: any[] = rawRes.themePreviews;
        if (themePreviews && themePreviews.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(themePreviews);
          }
          else {
            cmd.log(themePreviews.map(a => {
              return { Name: a.name };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No themes found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To get the list of available themes, you have to first connect to SharePoint
    using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    List available themes
      ${chalk.grey(config.delimiter)} ${commands.THEME_LIST}

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview
      `);
  }
}

module.exports = new SpoThemeListCommand();