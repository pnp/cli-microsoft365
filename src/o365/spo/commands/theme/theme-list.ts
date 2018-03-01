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

class ThemeListCommand extends SpoCommand {

  public get name(): string {
    return commands.THEME_LIST;
  }

  public get description(): string {
    return 'Retreives the list of custom themes added to the tenant store.';
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
              try {
                themePreviews.map(a => {
                  const themeJson = JSON.parse(a.themeJson);
                  cmd.log(`Name: ${a.name}\nPalette: ${JSON.stringify(themeJson.palette)}\n`);              
                });   
              }
              catch (e) 
              {
                cmd.log('Not able to retreive one of the theme.');
              }        
            }
          }
          else {
            if (this.verbose) {
              cmd.log('No themes found');
            }
          }        
        cb();     
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb);
      });
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
          
    Examples:
    
      Returns themes from the tenant store
      ${commands.THEME_LIST}
      
      Returns themes from the tenant store as JSON
      ${commands.THEME_LIST} -o json     
      `);
  }
}

module.exports = new ThemeListCommand();