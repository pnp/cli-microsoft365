import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import request from '../../../../request';
import { spo, ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { TermGroupCollection } from './TermGroupCollection';

class SpoTermGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_GROUP_LIST;
  }

  public get description(): string {
    return 'Lists taxonomy term groups';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoAdminUrl);
      if (this.verbose) {
        logger.logToStderr(`Retrieving taxonomy term groups...`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const result: TermGroupCollection = json[json.length - 1];
      if (result._Child_Items_ && result._Child_Items_.length > 0) {
        result._Child_Items_.forEach(t => {
          t.CreatedDate = new Date(Number(t.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
          t.Id = t.Id.replace('/Guid(', '').replace(')/', '');
          t.LastModifiedDate = new Date(Number(t.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        });
        logger.log(result._Child_Items_);
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTermGroupListCommand();