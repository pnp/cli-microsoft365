import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<{ "odata.null"?: boolean }> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.SPHSite/Details`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: { "odata.null"?: boolean }): void => {
        if (!res["odata.null"]) {
          logger.log(res);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoHomeSiteGetCommand();