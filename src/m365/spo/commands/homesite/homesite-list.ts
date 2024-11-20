import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHomeSiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_LIST;
  }

  public get description(): string {
    return 'Lists all home sites';
  }

  public defaultProperties(): string[] | undefined {
    return ['Url', 'Title'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      const res = await odata.getAllItems(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHomeSiteListCommand();