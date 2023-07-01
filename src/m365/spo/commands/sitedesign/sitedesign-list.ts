import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesign } from './SiteDesign.js';

class SpoSiteDesignListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_LIST;
  }

  public get description(): string {
    return 'Lists available site designs for creating modern sites';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'IsDefault', 'Title', 'Version', 'WebTemplate'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      const res: { value: SiteDesign[] } = await request.post(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignListCommand();