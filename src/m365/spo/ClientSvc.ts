import request from '../../request';
import Utils from '../../Utils';
import config from '../../config';
import { ClientSvcResponse, ClientSvcResponseContents } from './spo';
import { BasePermissions } from './base-permissions';
import { CommandInstance } from '../../cli';

export interface IdentityResponse {
  objectIdentity: string;
  serverRelativeUrl: string;
};

/**
 * Commonly used Client Svc calls.
 */
export class ClientSvc {
  constructor(private cmd: CommandInstance, private debug: boolean) {
    this.cmd = cmd;
    this.debug = debug;
  }

  /**
   * Requests web object identity for the current web.
   * That request is something similar to _contextinfo in REST.
   * The response data looks like:
   * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * _ObjectType_=SP.Web
   * ServerRelativeUrl=/sites/contoso
   * @param webUrl web url
   * @param formDigestValue formDigestValue
   */
  public getCurrentWebIdentity(webUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    return new Promise<IdentityResponse>((resolve: (identity: IdentityResponse) => void, reject: (error: any) => void): void => {
      request.post<string>(requestOptions).then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const identityObject: { _ObjectIdentity_: string; ServerRelativeUrl: string } = json.find(x => { return x._ObjectIdentity_; });
        if (identityObject) {
          return resolve({
            objectIdentity: identityObject._ObjectIdentity_,
            serverRelativeUrl: identityObject.ServerRelativeUrl
          });
        }

        reject('Cannot proceed. _ObjectIdentity_ not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); });
    });
  }

  /**
   * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
   * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * @param webUrl web url
   * @param siteAccessToken site access token
   * @param formDigestValue formDigestValue
   */
  public getEffectiveBasePermissions(webObjectIdentity: string, webUrl: string, formDigestValue: string): Promise<BasePermissions> {
    const basePermissionsResult: BasePermissions = new BasePermissions();

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    return new Promise<BasePermissions>((resolve: (permissions: BasePermissions) => void, reject: (error: any) => void): void => {
      request.post<string>(requestOptions).then((res: string): void => {
        if (this.debug) {
          this.cmd.log('Attempt to get the web EffectiveBasePermissions');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const permissionsObj = json.find(x => { return x.EffectiveBasePermissions; });
        if (permissionsObj) {
          basePermissionsResult.high = permissionsObj.EffectiveBasePermissions.High;
          basePermissionsResult.low = permissionsObj.EffectiveBasePermissions.Low;
          return resolve(basePermissionsResult);
        }

        reject('Cannot proceed. EffectiveBasePermissions not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); })
    });
  }

  /**
    * Gets folder by server relative url (GetFolderByServerRelativeUrl in REST)
    * The response data looks like:
    * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>:folder:<GUID>
    * _ObjectType_=SP.Folder
    * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
    * @param webUrl web url
    * @param siteRelativeUrl site relative url e.g. /Shared Documents/Folder1
    * @param formDigestValue formDigestValue
    */

  public getFolderIdentity(webObjectIdentity: string, webUrl: string, siteRelativeUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const serverRelativePath: string = Utils.getServerRelativePath(webUrl, siteRelativeUrl);

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativePath}</Parameter></Parameters></Method><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    return new Promise<IdentityResponse>((resolve: (identity: IdentityResponse) => void, reject: (error: any) => void) => {
      return request.post<string>(requestOptions).then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }
        const objectIdentity: { _ObjectIdentity_: string; } = json.find(x => { return x._ObjectIdentity_; });
        if (objectIdentity) {
          return resolve({
            objectIdentity: objectIdentity._ObjectIdentity_,
            serverRelativeUrl: serverRelativePath
          });
        }

        reject('Cannot proceed. Folder _ObjectIdentity_ not found'); // this is not suppose to happen
      }, (err: any): void => { reject(err); })
    });
  }
}