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

class ThemeGetCommand extends SpoCommand {

  public get name(): string {
    return commands.THEME_GET;
  }

  public get description(): string {
    return 'Gets the current theme settings for the site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Loading themes from tenant store...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving themes from tenant store...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/thememanager/GetTenantThemingOptions`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: '',
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }
        return request.get(requestOptions);
      })
      .then((themeResponse: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(themeResponse);
          cmd.log('');
        }

        cmd.log(themeResponse);
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
    
      To get information about a theme, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
          
    Example:
    
      Returns themes from the tenant store
      ${commands.THEME_GET}`);
  }
}

module.exports = new ThemeGetCommand();