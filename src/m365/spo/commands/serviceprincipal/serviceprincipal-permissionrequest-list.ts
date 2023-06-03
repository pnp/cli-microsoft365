import { OAuth2PermissionGrant, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import request, { CliRequestOptions } from '../../../../request';
import { ODataResponse } from '../../../../utils/odata';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SPOWebAppServicePrincipalPermissionRequest } from './SPOWebAppServicePrincipalPermissionRequest';

class SpoServicePrincipalPermissionRequestListCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_LIST;
  }

  public get description(): string {
    return 'Lists pending permission requests';
  }

  public alias(): string[] | undefined {
    return [commands.SP_PERMISSIONREQUEST_LIST];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      const reqDigest = await spo.getRequestDigest(spoAdminUrl);
      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><Query Id="13" ObjectPathId="11"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionRequests" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        let spoWebAppServicePrincipalPermissionRequestResult: SPOWebAppServicePrincipalPermissionRequest[] = [];

        const result: SPOWebAppServicePrincipalPermissionRequest[] = json[json.length - 1]._Child_Items_;
        if (result.length > 0) {
          const spoClientExtensibilityWebApplicationPrincipalId = await this.getSPOClientExtensibilityWebApplicationPrincipalId();
          if (spoClientExtensibilityWebApplicationPrincipalId !== null) {
            const oAuth2PermissionGrants: string[] | null = await this.getOAuth2PermissionGrants(spoClientExtensibilityWebApplicationPrincipalId);
            if (oAuth2PermissionGrants) {
              spoWebAppServicePrincipalPermissionRequestResult = result.filter(x => oAuth2PermissionGrants.indexOf(x.Scope) === -1);
            }
          }
        }
        if (spoWebAppServicePrincipalPermissionRequestResult.length === 0) {
          spoWebAppServicePrincipalPermissionRequestResult = result;
        }

        logger.log(spoWebAppServicePrincipalPermissionRequestResult.map(r => {
          return {
            Id: r.Id.replace('/Guid(', '').replace(')/', ''),
            Resource: r.Resource,
            ResourceId: r.ResourceId,
            Scope: r.Scope
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getOAuth2PermissionGrants(spoClientExtensibilityWebApplicationPrincipalId: string): Promise<string[] | null> {
    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/oAuth2Permissiongrants/?$filter=clientId eq '${spoClientExtensibilityWebApplicationPrincipalId}' and consentType eq 'AllPrincipals'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: ODataResponse<OAuth2PermissionGrant> = await request.get<ODataResponse<OAuth2PermissionGrant>>(requestOptions);
    if (response.value && response.value.length > 0) {
      return response.value[0].scope!.split(' ');
    }

    return null;
  }

  private async getSPOClientExtensibilityWebApplicationPrincipalId(): Promise<string | null> {
    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/servicePrincipals/?$filter=displayName eq 'SharePoint Online Client Extensibility Web Application Principal'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: ODataResponse<ServicePrincipal> = await request.get(requestOptions);
    if (response.value && response.value.length > 0) {
      return response.value[0].id!;
    }

    return null;
  }
}

module.exports = new SpoServicePrincipalPermissionRequestListCommand();