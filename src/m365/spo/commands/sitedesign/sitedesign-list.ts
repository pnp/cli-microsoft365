import { Logger } from '../../../../cli';
import request from '../../../../request';
import { spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesign } from './SiteDesign';

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
      logger.log(res.value);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignListCommand();