import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import request from '../../../../request';
import { spo, ClientSvcResponse, ClientSvcResponseContents } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoKnowledgehubGetCommand extends SpoCommand {
  public get name(): string {
    return commands.KNOWLEDGEHUB_GET;
  }

  public get description(): string {
    return 'Gets the Knowledge Hub Site URL for your tenant';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        logger.logToStderr(`Getting the Knowledge Hub Site settings for your tenant`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="5" ObjectPathId="4"/><Method Name="GetKnowledgeHubSite" Id="6" ObjectPathId="4"/></Actions><ObjectPaths><Constructor Id="4" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      
      const result: string = !json[json.length - 1] ? '' : json[json.length - 1];
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoKnowledgehubGetCommand();