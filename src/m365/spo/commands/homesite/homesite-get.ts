import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import { spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoHomeSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_GET;
  }

  public get description(): string {
    return 'Gets information about the Home Site';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${spoUrl}/_api/SP.SPHSite/Details`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ "odata.null"?: boolean }>(requestOptions);
      if (!res["odata.null"]) {
        logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoHomeSiteGetCommand();