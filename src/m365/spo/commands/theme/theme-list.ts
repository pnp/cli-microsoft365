import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((spoAdminUrl: string): Promise<any> => {
        if (this.verbose) {
          cmd.log(`Retrieving themes from tenant store...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/thememanager/GetTenantThemingOptions`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((rawRes: any): void => {
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
}

module.exports = new SpoThemeListCommand();