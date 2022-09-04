import { Logger } from '../../../../cli';
import config from '../../../../config';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { OrgAssets, OrgAssetsResponse } from './OrgAssets';

class SpoOrgAssetsLibraryListCommand extends SpoCommand {
  public get name(): string {
    return commands.ORGASSETSLIBRARY_LIST;
  }

  public get description(): string {
    return 'List all libraries that are assigned as asset library';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Method Name="GetOrgAssets" Id="6" ObjectPathId="3" /></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const orgAssetsResponse: OrgAssetsResponse = json[json.length - 1];

        if (orgAssetsResponse === null || orgAssetsResponse.OrgAssetsLibraries === undefined) {
          logger.log("No libraries in Organization Assets");
        }
        else {
          const orgAssets: OrgAssets = {
            Url: orgAssetsResponse.Url.DecodedUrl,
            Libraries: orgAssetsResponse.OrgAssetsLibraries._Child_Items_.map(t => {
              return {
                DisplayName: t.DisplayName,
                LibraryUrl: t.LibraryUrl.DecodedUrl,
                ListId: t.ListId,
                ThumbnailUrl: t.ThumbnailUrl !== null ? t.ThumbnailUrl.DecodedUrl : null
              };
            })
          };

          logger.log(orgAssets);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoOrgAssetsLibraryListCommand();
