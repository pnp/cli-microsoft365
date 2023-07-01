import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { ContextInfo, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
        await logger.log(res.value);
      }

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteScriptListCommand();