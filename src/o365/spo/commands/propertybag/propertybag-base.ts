import SpoCommand from "../../SpoCommand";
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { ClientSvcResponseContents, ClientSvcResponse } from "../../spo";
import config from '../../../../config';
import { BasePermissions } from "../../common/base-permissions";

export interface Property {
  key: string;
  value: any;
};

export interface IdentityResponse {
  objectIdentity: string;
  serverRelativeUrl: string;
};

export abstract class SpoPropertyBagBaseCommand extends SpoCommand {

  /**
   * Gets or sets site access token to be used 
   * with multiple methods.
   */
  protected siteAccessToken: string;

  /**
   * Gets or sets site form Digest Value to be used 
   * with multiple methods.
   */
  protected formDigestValue: string;

  /* istanbul ignore next */
  constructor() {
    super();
    this.siteAccessToken = '';
    this.formDigestValue = '';
  }

  /**
   * Requests web object itentity for the current web.
   * This request has to be send before we can construct the property bag request.
   * The response data looks like:
   * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * _ObjectType_=SP.Web
   * ServerRelativeUrl=/sites/contoso
   * The ObjectIdentity is needed to create another request to retrieve the property bag or set property.
   * @param webUrl web url
   * @param cmd command cmd
   */
  protected requestObjectIdentity(webUrl: string, cmd: CommandInstance): Promise<IdentityResponse> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    return new Promise<IdentityResponse>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');

          cmd.log('Attempt to get _ObjectIdentity_ key values');
        }

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const identityObject = json.find(x => { return x['_ObjectIdentity_'] });
        if (identityObject) {
          return resolve(
            {
              objectIdentity: identityObject['_ObjectIdentity_'],
              serverRelativeUrl: identityObject['ServerRelativeUrl']
            });
        }

        reject('Cannot proceed. _ObjectIdentity_ not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); });
    });
  }

  /**
   * This request has to be send before we can construct the property bag request.
   * The response data looks like:
   * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>:folder:<GUID>
   * _ObjectType_=SP.Folder
   * Properties ...
   * The ObjectIdentity is needed to create another request to retrieve the folder property bag or set property.
   * @param webUrl web url
   * @param cmd command cmd
   */
  protected requestFolderObjectIdentity(identityResp: IdentityResponse, webUrl: string, folder: string, cmd: CommandInstance): Promise<IdentityResponse> {
    let serverRelativeUrl: string = folder;
    if (identityResp.serverRelativeUrl !== '/') {
      serverRelativeUrl = `${identityResp.serverRelativeUrl}${serverRelativeUrl}`
    }

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativeUrl}</Parameter></Parameters></Method><Identity Id="5" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      cmd.log('Request:');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<IdentityResponse>((resolve: any, reject: any) => {

      return request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }
        const objectIdentity = json.find(x => { return x['_ObjectIdentity_'] });
        if (objectIdentity) {
          return resolve({
            objectIdentity: objectIdentity['_ObjectIdentity_'],
            serverRelativeUrl: serverRelativeUrl
          });
        }

        reject('Cannot proceed. Folder _ObjectIdentity_ not found'); // this is not suppose to happen

      }, (err: any): void => { reject(err); })
    });
  }

  /**
   * Gets property bag for a folder or site rootFolder of a site where return type is "_ObjectType_\":\"SP.Folder\".
   * This method is executed when folder option is specified. PnP PowerShell behaves the same way.
   */
  protected getFolderPropertyBag(identityResp: IdentityResponse, webUrl: string, folder: string, cmd: CommandInstance): Promise<Object> {
    let serverRelativeUrl: string = folder;
    if (identityResp.serverRelativeUrl !== '/') {
      serverRelativeUrl = `${identityResp.serverRelativeUrl}${serverRelativeUrl}`
    }

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativeUrl}</Parameter></Parameters></Method><Identity Id="5" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      cmd.log('Request:');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<Object>((resolve: any, reject: any) => {

      return request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');

          cmd.log('Attempt to get Properties key values');
        }

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const propertiesObj = json.find(x => { return x['Properties'] });
        if (propertiesObj) {
          return resolve(propertiesObj['Properties']);
        }

        reject('Cannot proceed. Properties not found'); // this is not suppose to happen

      }, (err: any): void => { reject(err); })
    });
  }

  /**
   * Gets property bag for site or sitecollection where return type is "_ObjectType_\":\"SP.Web\".
   * This method is executed when no folder specified. PnP PowerShell behaves the same way.
   */
  protected getWebPropertyBag(identityResp: IdentityResponse, webUrl: string, cmd: CommandInstance): Promise<Object> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="97" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /><Property Name="AllProperties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      cmd.log('Request:');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<Object>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');

          cmd.log('Attempt to get AllProperties key values');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const allPropertiesObj = json.find(x => { return x['AllProperties'] });
        if (allPropertiesObj) {
          return resolve(allPropertiesObj['AllProperties']);
        }

        reject('Cannot proceed. AllProperties not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); })
    });
  }

  /**
   * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
   * Note: This method can be moved as common method if is to be used for other commands.
   */
  protected getEffectiveBasePermissions(webObjectIdentity: string, webUrl: string, cmd: CommandInstance): Promise<BasePermissions> {

    let basePermissionsResult = new BasePermissions();

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      cmd.log('Request:');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<BasePermissions>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {

        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');

          cmd.log('Attempt to get the web EffectiveBasePermissions');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const permissionsObj = json.find(x => { return x['EffectiveBasePermissions'] });
        if (permissionsObj) {
          basePermissionsResult.high = permissionsObj['EffectiveBasePermissions']['High'];
          basePermissionsResult.low = permissionsObj['EffectiveBasePermissions']['Low'];
          return resolve(basePermissionsResult);
        }

        reject('Cannot proceed. EffectiveBasePermissions not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); })
    });
  }

  /**
   * The property bag item data returned from the client.svc/ProcessQuery response
   * has to be formatted before displayed since the key, value objects
   * carry extra information or there might be a value,
   * that should to be formatted depending on the data type.
   */
  protected formatProperty(objKey: string, objValue: any): Property {
    if (objKey.indexOf('$  Int32') > -1) {

      // format if the propery value is integer
      // the int returned has the following format of the property key,
      // 'vti_folderitemcount$  Int32'. To normalize that, the extra string
      // '$  Int32' has to be removed from the key, also parseInt is used to 
      // ensure the json object returns number

      objKey = objKey.replace('$  Int32', '');
      objValue = parseInt(objValue);
    }
    else {
      if (typeof objValue === 'string' && objValue.indexOf('/Date(') > -1) {

        // format if the property value is date
        // the date returned has the following format ex. /Date(2017,10,7,11,29,31,0)/.
        // That has to be turned into JavaScript Date object

        let date = objValue.replace('/Date(', '').replace(')/', '').split(',').map(Number);
        objValue = new Date(date[0], date[1], date[2], date[3], date[4], date[5], date[6]);
      }
      else {
        if (objValue === 'true' || objValue === 'false') {

          // format if the property value is boolean
          objValue = (objValue === 'true');
        }
      }
    }

    return { key: objKey, value: objValue } as Property;
  }
}