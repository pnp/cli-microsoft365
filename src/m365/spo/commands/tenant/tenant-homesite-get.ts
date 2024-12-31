import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoTenantHomeSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_HOMESITE_GET;
  }

  public get description(): string {
    return 'Gets information about the Home Site';
  }

  public alias(): string[] {
    return ['spo homesite get'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()[0], this.getCommandName());
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
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantHomeSiteGetCommand();