import { Logger } from '../../../../cli';
import config from '../../../../config';
import request from '../../../../request';
import { spo, ClientSvcResponse, ClientSvcResponseContents } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoContentTypeHubGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPEHUB_GET;
  }

  public get description(): string {
    return 'Returns the URL of the SharePoint Content Type Hub of the Tenant';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoUrl = await spo.getSpoUrl(logger,this.debug);
      const reqDigest = await spo.getRequestDigest(spoUrl);

      if (this.verbose) {
        logger.logToStderr(`Retrieving Content Type Hub URL`);
      }

      const requestOptions: any = {
        url: `${spoUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}">
<Actions>
  <ObjectPath Id="2" ObjectPathId="1" />
  <ObjectIdentityQuery Id="3" ObjectPathId="1" />
  <ObjectPath Id="5" ObjectPathId="4" />
  <ObjectIdentityQuery Id="6" ObjectPathId="4" />
  <Query Id="7" ObjectPathId="4">
    <Query SelectAllProperties="false">
      <Properties>
        <Property Name="ContentTypePublishingHub" ScalarProperty="true" />
      </Properties>
    </Query>
  </Query>
</Actions>
<ObjectPaths>
  <StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
  <Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" />
</ObjectPaths>
</Request>`
      };

      const res = await request.post<string>(requestOptions);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      const result: any = {
        ContentTypePublishingHub: json[json.length - 1]["ContentTypePublishingHub"]
      }; 
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoContentTypeHubGetCommand();