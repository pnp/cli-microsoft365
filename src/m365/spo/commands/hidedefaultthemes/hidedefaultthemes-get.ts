import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import { CommandInstance } from '../../../../cli';

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
}

module.exports = new SpoHideDefaultThemesGetCommand();