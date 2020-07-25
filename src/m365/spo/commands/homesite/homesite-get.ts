import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class SpoHomeSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_GET;
  }

  public get description(): string {
    return 'Gets information about the Home Site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<{ "odata.null"?: boolean }> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.SPHSite/Details`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: { "odata.null"?: boolean }): void => {
        if (!res["odata.null"]) {
          cmd.log(res);
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new SpoHomeSiteGetCommand();