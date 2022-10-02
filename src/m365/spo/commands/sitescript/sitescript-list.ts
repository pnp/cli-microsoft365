import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import { spo, ContextInfo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoSiteScriptListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_LIST;
  }

  public get description(): string {
    return 'Lists site script available for use with site designs';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const formDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts'`,
        headers: {
          'X-RequestDigest': formDigest.FormDigestValue,
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.post<{ value: any[] }>(requestOptions);
      if (res.value && res.value.length > 0) {
        logger.log(res.value);
      }
      
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteScriptListCommand();