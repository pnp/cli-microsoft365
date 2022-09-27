import { Logger } from '../../../../cli';
import config from '../../../../config';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoOrgNewsSiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.ORGNEWSSITE_LIST;
  }

  public get description(): string {
    return 'Lists all organizational news sites';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const resDigest = await spo.getRequestDigest(spoAdminUrl);

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': resDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="58" ObjectPathId="57" /><Method Name="GetOrgNewsSites" Id="59" ObjectPathId="57" /></Actions><ObjectPaths><Constructor Id="57" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const results: string[] = json[json.length - 1];
        logger.log(results);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoOrgNewsSiteListCommand();