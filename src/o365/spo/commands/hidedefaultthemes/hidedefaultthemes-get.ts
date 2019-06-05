import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class SpoHideDefaultThemesGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HIDEDEFAULTTHEMES_GET;
  }

  public get description(): string {
    return 'Gets the current value of the HideDefaultThemes setting';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((spoAdminUrl: string): Promise<any> => {
        if (this.verbose) {
          cmd.log(`Getting the current value of the HideDefaultThemes setting...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/thememanager/GetHideDefaultThemes`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((rawRes: any): void => {
        cmd.log(rawRes.value);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
      
  Examples:
  
    Get the current value of the HideDefaultThemes setting
      ${commands.HIDEDEFAULTTHEMES_GET}

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview
      `);
  }
}

module.exports = new SpoHideDefaultThemesGetCommand();