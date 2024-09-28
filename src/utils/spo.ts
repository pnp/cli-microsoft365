import os from 'os';
import url from 'url';
import { urlUtil } from "./urlUtil.js";
import { validation } from "./validation.js";
import auth from '../Auth.js';
import { Logger } from "../cli/Logger.js";
import config from "../config.js";
import { BasePermissions } from '../m365/spo/base-permissions.js';
import request, { CliRequestOptions } from "../request.js";
import { formatting } from './formatting.js';
import { CustomAction } from '../m365/spo/commands/customaction/customaction.js';
import { MenuState } from '../m365/spo/commands/navigation/NavigationNode.js';
import { odata } from './odata.js';
import { RoleDefinition } from '../m365/spo/commands/roledefinition/RoleDefinition.js';
import { RoleType } from '../m365/spo/commands/roledefinition/RoleType.js';
import { DeletedSiteProperties } from '../m365/spo/commands/site/DeletedSiteProperties.js';
import { SiteProperties } from '../m365/spo/commands/site/SiteProperties.js';
import { entraGroup } from './entraGroup.js';
import { SharingCapabilities } from '../m365/spo/commands/site/SharingCapabilities.js';
import { WebProperties } from '../m365/spo/commands/web/WebProperties.js';
import { Group, Site } from '@microsoft/microsoft-graph-types';
import { ListItemInstance } from '../m365/spo/commands/listitem/ListItemInstance.js';
import { ListItemFieldValueResult } from '../m365/spo/commands/listitem/ListItemFieldValueResult.js';
import { FileProperties } from '../m365/spo/commands/file/FileProperties.js'; import { setTimeout } from 'timers/promises';

export interface ContextInfo {
  FormDigestTimeoutSeconds: number;
  FormDigestValue: string;
  WebFullUrl: string;
}

export interface FormDigestInfo extends ContextInfo {
  FormDigestExpiresAt: Date;
}

export type ClientSvcResponse = Array<any | ClientSvcResponseContents>;

export interface ClientSvcResponseContents {
  SchemaVersion: string;
  LibraryVersion: string;
  ErrorInfo?: {
    ErrorMessage: string;
    ErrorValue?: string;
    TraceCorrelationId: string;
    ErrorCode: number;
    ErrorTypeName?: string;
  };
  TraceCorrelationId: string;
}

export interface SpoOperation {
  _ObjectIdentity_: string;
  IsComplete: boolean;
  PollingInterval: number;
}

export interface IdentityResponse {
  objectIdentity: string;
  serverRelativeUrl: string;
}

export interface GraphFileDetails {
  SiteId: string;
  VroomDriveID: string;
  VroomItemID: string;
}

interface FormValues {
  FieldName: string;
  FieldValue: string;
}

export interface User {
  Id: number;
  IsHiddenInUI: boolean;
  LoginName: string;
  Title: string;
  PrincipalType: number;
  Email: string;
  Expiration: string;
  IsEmailAuthenticationGuestUser: boolean;
  IsShareByEmailGuestUser: boolean;
  IsSiteAdmin: boolean;
  UserId: {
    NameId: string;
    NameIdIssuer: string;
  } | null;
  UserPrincipalName: string | null;
}

export const spo = {
  async getRequestDigest(siteUrl: string): Promise<FormDigestInfo> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  },

  async ensureFormDigest(siteUrl: string, logger: Logger, context: FormDigestInfo | undefined, debug: boolean): Promise<FormDigestInfo> {
    if (validation.isValidFormDigest(context)) {
      if (debug) {
        await logger.logToStderr('Existing form digest still valid');
      }

      return context as FormDigestInfo;
    }

    const res: FormDigestInfo = await spo.getRequestDigest(siteUrl);
    const now: Date = new Date();
    now.setSeconds(now.getSeconds() + res.FormDigestTimeoutSeconds - 5);
    context = {
      FormDigestValue: res.FormDigestValue,
      FormDigestTimeoutSeconds: res.FormDigestTimeoutSeconds,
      FormDigestExpiresAt: now,
      WebFullUrl: res.WebFullUrl
    };
    return context;
  },

  async waitUntilFinished({ operationId, siteUrl, logger, currentContext, debug, verbose }: { operationId: string, siteUrl: string, logger: Logger, currentContext: FormDigestInfo, debug: boolean, verbose: boolean }): Promise<void> {
    const resFormDigest = await spo.ensureFormDigest(siteUrl, logger, currentContext, debug);
    currentContext = resFormDigest;
    if (debug) {
      await logger.logToStderr(`Checking if operation ${operationId} completed...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': currentContext.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];
    if (response.ErrorInfo) {
      throw new Error(response.ErrorInfo.ErrorMessage);
    }
    else {
      const operation: SpoOperation = json[json.length - 1];
      const isComplete: boolean = operation.IsComplete;
      if (isComplete) {
        if (!debug && verbose) {
          process.stdout.write('\n');
        }

        return;
      }

      await setTimeout(operation.PollingInterval);
      await spo.waitUntilFinished({
        operationId: JSON.stringify(operation._ObjectIdentity_),
        siteUrl,
        logger,
        currentContext,
        debug,
        verbose
      });
    }
  },

  async getSpoUrl(logger: Logger, debug: boolean): Promise<string> {
    if (auth.connection.spoUrl) {
      if (debug) {
        await logger.logToStderr(`SPO URL previously retrieved ${auth.connection.spoUrl}. Returning...`);
      }

      return auth.connection.spoUrl;
    }

    if (debug) {
      await logger.logToStderr(`No SPO URL available. Retrieving from MS Graph...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res: { webUrl: string } = await request.get<{ webUrl: string }>(requestOptions);

    auth.connection.spoUrl = res.webUrl;
    try {
      await auth.storeConnectionInfo();
    }
    catch (e: any) {
      if (debug) {
        await logger.logToStderr('Error while storing connection info');
      }
    }
    return auth.connection.spoUrl;
  },

  async getSpoAdminUrl(logger: Logger, debug: boolean): Promise<string> {
    const spoUrl = await spo.getSpoUrl(logger, debug);
    return (spoUrl.replace(/(https:\/\/)([^\.]+)(.*)/, '$1$2-admin$3'));
  },

  async getTenantId(logger: Logger, debug: boolean): Promise<string> {
    if (auth.connection.spoTenantId) {
      if (debug) {
        await logger.logToStderr(`SPO Tenant ID previously retrieved ${auth.connection.spoTenantId}. Returning...`);
      }

      return auth.connection.spoTenantId;
    }

    if (debug) {
      await logger.logToStderr(`No SPO Tenant ID available. Retrieving...`);
    }

    const spoAdminUrl = await spo.getSpoAdminUrl(logger, debug);
    const contextInfo: ContextInfo = await spo.getRequestDigest(spoAdminUrl);

    const tenantInfoRequestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': contextInfo.FormDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post(tenantInfoRequestOptions);
    const json: string[] = JSON.parse(res);
    auth.connection.spoTenantId = (json[json.length - 1] as any)._ObjectIdentity_.replace('\n', '&#xA;');
    try {
      await auth.storeConnectionInfo();
    }
    catch (e: any) {
      if (debug) {
        await logger.logToStderr('Error while storing connection info');
      }
    }
    return auth.connection.spoTenantId as string;
  },

  /**
   * Returns the Graph id of a site 
   * @param webUrl web url e.g. https://contoso.sharepoint.com/sites/site1
   */
  async getSpoGraphSiteId(webUrl: string): Promise<string> {
    const url = new URL(webUrl);

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}?$select=id`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const result = await request.get<{ id: string }>(requestOptions);
    return result.id;
  },

  /**
   * Ensures the folder path exists
   * @param webFullUrl web full url e.g. https://contoso.sharepoint.com/sites/site1
   * @param folderToEnsure web relative or server relative folder path e.g. /Documents/MyFolder or /sites/site1/Documents/MyFolder
   * @param siteAccessToken a valid access token for the site specified in the webFullUrl param
   */
  async ensureFolder(webFullUrl: string, folderToEnsure: string, logger: Logger, debug: boolean): Promise<void> {
    const webUrl = url.parse(webFullUrl);
    if (!webUrl.protocol || !webUrl.hostname) {
      throw new Error('webFullUrl is not a valid URL');
    }

    if (!folderToEnsure) {
      throw new Error('folderToEnsure cannot be empty');
    }

    // remove last '/' of webFullUrl if exists
    const webFullUrlLastCharPos: number = webFullUrl.length - 1;

    if (webFullUrl.length > 1 &&
      webFullUrl[webFullUrlLastCharPos] === '/') {
      webFullUrl = webFullUrl.substring(0, webFullUrlLastCharPos);
    }

    folderToEnsure = urlUtil.getWebRelativePath(webFullUrl, folderToEnsure);

    if (debug) {
      await logger.log(`folderToEnsure`);
      await logger.log(folderToEnsure);
      await logger.log('');
    }

    let nextFolder: string = '';
    let prevFolder: string = '';
    let folderIndex: number = 0;

    // build array of folders e.g. ["Shared%20Documents","22","54","55"]
    const folders: string[] = folderToEnsure.substring(1).split('/');

    if (debug) {
      await logger.log('folders to process');
      await logger.log(JSON.stringify(folders));
      await logger.log('');
    }

    // recursive function
    async function checkOrAddFolder(): Promise<void> {
      if (folderIndex === folders.length) {
        if (debug) {
          await logger.log(`All sub-folders exist`);
        }

        return;
      }

      // append the next sub-folder to the folder path and check if it exists
      prevFolder = nextFolder;
      nextFolder += `/${folders[folderIndex]}`;
      const folderServerRelativeUrl = urlUtil.getServerRelativePath(webFullUrl, nextFolder);

      const requestOptions: CliRequestOptions = {
        url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        }
      };
      try {
        await request.get(requestOptions);

        folderIndex++;
        await checkOrAddFolder();
      }
      catch {
        const prevFolderServerRelativeUrl: string = urlUtil.getServerRelativePath(webFullUrl, prevFolder);
        const requestOptions: CliRequestOptions = {
          url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27${formatting.encodeQueryParameter(prevFolderServerRelativeUrl)}%27&@a2=%27${formatting.encodeQueryParameter(folders[folderIndex])}%27`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        try {
          await request.post(requestOptions);
          folderIndex++;
          await checkOrAddFolder();
        }
        catch (err) {
          if (debug) {
            await logger.log(`Could not create sub-folder ${folderServerRelativeUrl}`);
          }
          throw err;
        }


      }
    };

    return checkOrAddFolder();
  },

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
  async getCurrentWebIdentity(webUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    const res: string = await request.post<string>(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);

    const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
    if (contents && contents.ErrorInfo) {
      throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
    }

    const identityObject: { _ObjectIdentity_: string; ServerRelativeUrl: string } = json.find(x => { return x._ObjectIdentity_; });
    if (identityObject) {
      return {
        objectIdentity: identityObject._ObjectIdentity_,
        serverRelativeUrl: identityObject.ServerRelativeUrl
      };
    }
    throw 'Cannot proceed. _ObjectIdentity_ not found';
  },

  /**
   * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
   * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * @param webUrl web url
   * @param siteAccessToken site access token
   * @param formDigestValue formDigestValue
   */
  async getEffectiveBasePermissions(webObjectIdentity: string, webUrl: string, formDigestValue: string, logger: Logger, debug: boolean): Promise<BasePermissions> {
    const basePermissionsResult: BasePermissions = new BasePermissions();

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };


    const res: string = await request.post<string>(requestOptions);
    if (debug) {
      await logger.log('Attempt to get the web EffectiveBasePermissions');
    }

    const json: ClientSvcResponse = JSON.parse(res);
    const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
    if (contents && contents.ErrorInfo) {
      throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
    }

    const permissionsObj = json.find(x => { return x.EffectiveBasePermissions; });
    if (permissionsObj) {
      basePermissionsResult.high = permissionsObj.EffectiveBasePermissions.High;
      basePermissionsResult.low = permissionsObj.EffectiveBasePermissions.Low;
      return basePermissionsResult;
    }

    throw ('Cannot proceed. EffectiveBasePermissions not found'); // this is not supposed to happen
  },

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
  async getFolderIdentity(webObjectIdentity: string, webUrl: string, siteRelativeUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const serverRelativePath: string = urlUtil.getServerRelativePath(webUrl, siteRelativeUrl);

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativePath}</Parameter></Parameters></Method><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post<string>(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);

    const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
    if (contents && contents.ErrorInfo) {
      throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
    }
    const objectIdentity: { _ObjectIdentity_: string; } = json.find(x => { return x._ObjectIdentity_; });
    if (objectIdentity) {
      return {
        objectIdentity: objectIdentity._ObjectIdentity_,
        serverRelativeUrl: serverRelativePath
      };
    }

    throw 'Cannot proceed. Folder _ObjectIdentity_ not found';
  },

  /**
   * Retrieves the SiteId, VroomItemId and VroomDriveId from a specific file.
   * @param webUrl Web url
   * @param fileId GUID ID of the file
   * @param fileUrl Decoded site-relative or server-relative URL of the file
   */
  async getVroomFileDetails(webUrl: string, fileId?: string, fileUrl?: string): Promise<GraphFileDetails> {
    let requestUrl: string = `${webUrl}/_api/web/`;

    if (fileUrl) {
      const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')`;
    }
    else {
      requestUrl += `GetFileById('${fileId}')`;
    }

    requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId';

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<GraphFileDetails>(requestOptions);
    return res;
  },

  /**
   * Retrieves a list of Custom Actions from a SharePoint site.
   * @param webUrl Web url
   * @param scope The scope of custom actions to retrieve, allowed values "Site", "Web" or "All".
   * @param filter An OData filter query to limit the results.
   */
  async getCustomActions(webUrl: string, scope: string | undefined, filter?: string): Promise<CustomAction[]> {
    if (scope && scope !== "All" && scope !== "Site" && scope !== "Web") {
      throw `Invalid scope '${scope}'. Allowed values are 'Site', 'Web' or 'All'.`;
    }

    const queryString = filter ? `?$filter=${filter}` : "";

    if (scope && scope !== "All") {
      return await odata.getAllItems<CustomAction>(`${webUrl}/_api/${scope}/UserCustomActions${queryString}`);
    }

    const customActions = [
      ...await odata.getAllItems<CustomAction>(`${webUrl}/_api/Site/UserCustomActions${queryString}`),
      ...await odata.getAllItems<CustomAction>(`${webUrl}/_api/Web/UserCustomActions${queryString}`)
    ];

    return customActions;
  },


  /**
   * Retrieves a Custom Actions from a SharePoint site by Id.
   * @param webUrl Web url
   * @param id The Id of the Custom Action
   * @param scope The scope of custom actions to retrieve, allowed values "Site", "Web" or "All".
   */
  async getCustomActionById(webUrl: string, id: string, scope?: string): Promise<CustomAction | undefined> {
    if (scope && scope !== "All" && scope !== "Site" && scope !== "Web") {
      throw `Invalid scope '${scope}'. Allowed values are 'Site', 'Web' or 'All'.`;
    }

    async function getById(webUrl: string, id: string, scope: string): Promise<CustomAction | undefined> {
      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/${scope}/UserCustomActions(guid'${id}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const result = await request.get<CustomAction>(requestOptions);

      if (result["odata.null"] === true) {
        return undefined;
      }

      return result;
    }

    if (scope && scope !== "All") {
      return await getById(webUrl, id, scope);
    }

    const customActionOnWeb = await getById(webUrl, id, "Web");
    if (customActionOnWeb) {
      return customActionOnWeb;
    }

    const customActionOnSite = await getById(webUrl, id, "Site");
    return customActionOnSite;
  },

  async getTenantAppCatalogUrl(logger: Logger, debug: boolean): Promise<string | null> {
    const spoUrl = await spo.getSpoUrl(logger, debug);

    const requestOptions: CliRequestOptions = {
      url: `${spoUrl}/_api/SP_TenantSettings_Current`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const result = await request.get<{ CorporateCatalogUrl: string }>(requestOptions);
    return result.CorporateCatalogUrl;
  },

  /**
   * Retrieves the Microsoft Entra ID from a SP user.
   * @param webUrl Web url
   * @param id The Id of the user
   */
  async getUserAzureIdBySpoId(webUrl: string, id: string): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/siteusers/GetById('${formatting.encodeQueryParameter(id)}')?$select=AadObjectId`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<{ AadObjectId: { NameId: string, NameIdIssuer: string } }>(requestOptions);

    return res.AadObjectId.NameId;
  },

  /**
   * Ensure a user exists on a specific SharePoint site.
   * @param webUrl URL of the SharePoint site.
   * @param logonName Logon name of the user to ensure on the SharePoint site.
   * @returns SharePoint user object.
   */
  async ensureUser(webUrl: string, logonName: string): Promise<User> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/EnsureUser`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        logonName: logonName
      }
    };

    return request.post<User>(requestOptions);
  },

  /**
   * Ensure a Microsoft Entra ID group exists on a specific SharePoint site.
   * @param webUrl URL of the SharePoint site.
   * @param group Microsoft Entra ID group.
   * @returns SharePoint user object.
   */
  async ensureEntraGroup(webUrl: string, group: Group): Promise<User> {
    if (!group.securityEnabled) {
      throw new Error('Cannot ensure a Microsoft Entra ID group that is not security enabled.');
    }

    return this.ensureUser(webUrl, group.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${group.id}` : `c:0t.c|tenant|${group.id}`);
  },

  /**
 * Retrieves the spo user by email.
 * @param webUrl Web url
 * @param email The email of the user
 * @param logger the Logger object
 * @param verbose set for verbose logging
 */
  async getUserByEmail(webUrl: string, email: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving the spo user by email ${email}`);
    }
    const requestUrl = `${webUrl}/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter(email)}')`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get<User>(requestOptions);

    return userInstance;
  },

  /**
   * Retrieves the menu state for the quick launch.
   * @param webUrl Web url
   */
  async getQuickLaunchMenuState(webUrl: string): Promise<MenuState> {
    return this.getMenuState(webUrl);
  },

  /**
   * Retrieves the menu state for the top navigation.
   * @param webUrl Web url
   */
  async getTopNavigationMenuState(webUrl: string): Promise<MenuState> {
    return this.getMenuState(webUrl, '1002');
  },

  /**
   * Retrieves the menu state.
   * @param webUrl Web url
   * @param menuNodeKey Menu node key
   */
  async getMenuState(webUrl: string, menuNodeKey?: string): Promise<MenuState> {
    const requestBody = {
      customProperties: null,
      depth: 10,
      mapProviderName: null,
      menuNodeKey: menuNodeKey || null
    };
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/navigation/MenuState`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    return request.post<MenuState>(requestOptions);
  },

  /**
  * Saves the menu state.
  * @param webUrl Web url
  * @param menuState Updated menu state
  */
  async saveMenuState(webUrl: string, menuState: MenuState): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/navigation/SaveMenuState`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: { menuState: menuState },
      responseType: 'json'
    };

    return request.post(requestOptions);
  },

  /**
  * Retrieves the spo group by name.
  * @param webUrl Web url
  * @param name The name of the group
  * @param logger the Logger object
  * @param verbose set for verbose logging
  */
  async getGroupByName(webUrl: string, name: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving the group by name ${name}`);
    }
    const requestUrl = `${webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(name)}')`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const groupInstance: any = await request.get(requestOptions);

    return groupInstance;
  },

  /**
  * Retrieves the role definition by name.
  * @param webUrl Web url
  * @param name the name of the role definition
  * @param logger the Logger object
  * @param verbose set for verbose logging
  */
  async getRoleDefinitionByName(webUrl: string, name: string, logger?: Logger, verbose?: boolean): Promise<RoleDefinition> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving the role definitions for ${name}`);
    }

    const roledefinitions = await odata.getAllItems<RoleDefinition>(`${webUrl}/_api/web/roledefinitions`);
    const roledefinition = roledefinitions.find((role: RoleDefinition) => role.Name === name);

    if (!roledefinition) {
      throw `No roledefinition is found for ${name}`;
    }

    const permissions: BasePermissions = new BasePermissions();
    permissions.high = roledefinition.BasePermissions.High as number;
    permissions.low = roledefinition.BasePermissions.Low as number;
    roledefinition.BasePermissionsValue = permissions.parse();
    roledefinition.RoleTypeKindValue = RoleType[roledefinition.RoleTypeKind];

    return roledefinition;
  },

  /**
  * Adds a SharePoint site.
  * @param type Type of sites to add. Allowed values TeamSite, CommunicationSite, ClassicSite, default TeamSite
  * @param title Site title
  * @param alias Site alias, used in the URL and in the team site group e-mail (applies to type TeamSite)
  * @param url Site URL (applies to type CommunicationSite, ClassicSite)
  * @param timeZone Integer representing time zone to use for the site (applies to type ClassicSite)
  * @param description Site description
  * @param lcid Site language in the LCID format
  * @param owners Comma-separated list of users to set as site owners (applies to type TeamSite, ClassicSite)
  * @param isPublic Determines if the associated group is public or not (applies to type TeamSite)
  * @param classification Site classification (applies to type TeamSite, CommunicationSite)
  * @param siteDesignType of communication site to create. Allowed values Topic, Showcase, Blank, default Topic. When creating a communication site, specify either siteDesign or siteDesignId (applies to type CommunicationSite)
  * @param siteDesignId Id of the custom site design to use to create the site. When creating a communication site, specify either siteDesign or siteDesignId (applies to type CommunicationSite)
  * @param shareByEmailEnabled Determines whether it's allowed to share file with guests (applies to type CommunicationSite)
  * @param webTemplate Template to use for creating the site. Default `STS#0` (applies to type ClassicSite)
  * @param resourceQuota The quota for this site collection in Sandboxed Solutions units. Default 0 (applies to type ClassicSite)
  * @param resourceQuotaWarningLevel The warning level for the resource quota. Default 0 (applies to type ClassicSite)
  * @param storageQuota The storage quota for this site collection in megabytes. Default 100 (applies to type ClassicSite)
  * @param storageQuotaWarningLevel The warning level for the storage quota in megabytes. Default 100 (applies to type ClassicSite)
  * @param removeDeletedSite Set, to remove existing deleted site with the same URL from the Recycle Bin (applies to type ClassicSite)
  * @param wait Wait for the site to be provisioned before completing the command (applies to type ClassicSite)
  * @param logger the Logger object
  * @param verbose set if verbose logging should be logged 
  */
  async addSite(title: string, logger: Logger, verbose: boolean, wait?: boolean, type?: string, alias?: string, description?: string, owners?: string, shareByEmailEnabled?: boolean, removeDeletedSite?: boolean, classification?: string, isPublic?: boolean, lcid?: number, url?: string, siteDesign?: string, siteDesignId?: string, timeZone?: string | number, webTemplate?: string, resourceQuota?: string | number, resourceQuotaWarningLevel?: string | number, storageQuota?: string | number, storageQuotaWarningLevel?: string | number): Promise<any> {
    interface CreateGroupExResponse {
      DocumentsUrl: string;
      ErrorMessage: string;
      GroupId: string;
      SiteStatus: number;
      SiteUrl: string;
    }

    if (type === 'ClassicSite') {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
      let context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);

      let exists: boolean;
      if (removeDeletedSite) {
        exists = await spo.siteExists(url as string, logger, verbose);
      }
      else {
        // assume site doesn't exist
        exists = false;
      }

      if (exists) {
        if (verbose) {
          await logger.logToStderr('Site exists in the recycle bin');
        }

        await spo.deleteSiteFromTheRecycleBin(url as string, logger, verbose, wait);
      }
      else {
        if (verbose) {
          await logger.logToStderr('Site not found');
        }
      }

      context = await spo.ensureFormDigest(spoAdminUrl as string, logger, context, verbose);

      if (verbose) {
        await logger.logToStderr(`Creating site collection ${url}...`);
      }

      const lcidOption: number = typeof lcid === 'number' ? lcid : 1033;
      const storageQuotaOption: number = typeof storageQuota === 'number' ? storageQuota : 100;
      const storageQuotaWarningLevelOption: number = typeof storageQuotaWarningLevel === 'number' ? storageQuotaWarningLevel : 100;
      const resourceQuotaOption: number = typeof resourceQuota === 'number' ? resourceQuota : 0;
      const resourceQuotaWarningLevelOption: number = typeof resourceQuotaWarningLevel === 'number' ? resourceQuotaWarningLevel : 0;
      const webTemplateOption: string = webTemplate || 'STS#0';

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">${lcidOption}</Property><Property Name="Owner" Type="String">${formatting.escapeXml(owners)}</Property><Property Name="StorageMaximumLevel" Type="Int64">${storageQuotaOption}</Property><Property Name="StorageWarningLevel" Type="Int64">${storageQuotaWarningLevelOption}</Property><Property Name="Template" Type="String">${formatting.escapeXml(webTemplateOption)}</Property><Property Name="TimeZoneId" Type="Int32">${timeZone}</Property><Property Name="Title" Type="String">${formatting.escapeXml(title)}</Property><Property Name="Url" Type="String">${formatting.escapeXml(url)}</Property><Property Name="UserCodeMaximumLevel" Type="Double">${resourceQuotaOption}</Property><Property Name="UserCodeWarningLevel" Type="Double">${resourceQuotaWarningLevelOption}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const operation: SpoOperation = json[json.length - 1];
        const isComplete: boolean = operation.IsComplete;
        if (!wait || isComplete) {
          return;
        }

        await setTimeout(operation.PollingInterval);
        await spo.waitUntilFinished({
          operationId: JSON.stringify(operation._ObjectIdentity_),
          siteUrl: spoAdminUrl,
          logger,
          currentContext: context,
          verbose: verbose,
          debug: verbose
        });
      }
    }
    else {
      const isTeamSite: boolean = type !== 'CommunicationSite';

      const spoUrl = await spo.getSpoUrl(logger, verbose);

      if (verbose) {
        await logger.logToStderr(`Creating new site...`);
      }

      let requestOptions: any = {};

      if (isTeamSite) {
        requestOptions = {
          url: `${spoUrl}/_api/GroupSiteManager/CreateGroupEx`,
          headers: {
            'content-type': 'application/json; odata=verbose; charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            displayName: title,
            alias: alias,
            isPublic: isPublic,
            optionalParams: {
              Description: description || '',
              CreationOptions: {
                results: [],
                Classification: classification || ''
              }
            }
          }
        };

        if (lcid) {
          requestOptions.data.optionalParams.CreationOptions.results.push(`SPSiteLanguage:${lcid}`);
        }

        if (owners) {
          requestOptions.data.optionalParams.Owners = {
            results: owners.split(',').map(o => o.trim())
          };
        }
      }
      else {
        if (siteDesignId) {
          siteDesignId = siteDesignId;
        }
        else {
          if (siteDesign) {
            switch (siteDesign) {
              case 'Topic':
                siteDesignId = '00000000-0000-0000-0000-000000000000';
                break;
              case 'Showcase':
                siteDesignId = '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
                break;
              case 'Blank':
                siteDesignId = 'f6cc5403-0d63-442e-96c0-285923709ffc';
                break;
            }
          }
          else {
            siteDesignId = '00000000-0000-0000-0000-000000000000';
          }
        }

        requestOptions = {
          url: `${spoUrl}/_api/SPSiteManager/Create`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            request: {
              Title: title,
              Url: url,
              ShareByEmailEnabled: shareByEmailEnabled,
              Description: description || '',
              Classification: classification || '',
              WebTemplate: 'SITEPAGEPUBLISHING#0',
              SiteDesignId: siteDesignId
            }
          }
        };

        if (lcid) {
          requestOptions.data.request.Lcid = lcid;
        }

        if (owners) {
          requestOptions.data.request.Owner = owners;
        }
      }

      const res = await request.post<CreateGroupExResponse>(requestOptions);

      if (isTeamSite) {
        if (res.ErrorMessage !== null) {
          throw res.ErrorMessage;
        }
        else {
          return res.SiteUrl;
        }
      }
      else {
        if (res.SiteStatus === 2) {
          return res.SiteUrl;
        }
        else {
          throw 'An error has occurred while creating the site';
        }
      }
    }
  },

  /**
  * Checks if a site exists
  * Returns a boolean
  * @param url The url of the site
  * @param logger the Logger object
  * @param verbose set if verbose logging should be logged 
  */
  async siteExists(url: string, logger: Logger, verbose: boolean): Promise<boolean> {
    const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
    const context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);

    if (verbose) {
      await logger.logToStderr(`Checking if the site ${url} exists...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': context.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`
    };

    const res1: any = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res1);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      if (response.ErrorInfo.ErrorTypeName === 'Microsoft.Online.SharePoint.Common.SpoNoSiteException') {
        return await this.siteExistsInTheRecycleBin(url, logger, verbose);
      }
      else {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    else {
      const site: SiteProperties = json[json.length - 1];

      return site.Status === 'Recycled';
    }
  },

  /**
  * Checks if a site exists in the recycle bin
  * Returns a boolean
  * @param url The url of the site
  * @param logger the Logger object
  * @param verbose set if verbose logging should be logged 
  */
  async siteExistsInTheRecycleBin(url: string, logger: Logger, verbose: boolean): Promise<boolean> {
    if (verbose) {
      await logger.logToStderr(`Site doesn't exist. Checking if the site ${url} exists in the recycle bin...`);
    }

    const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
    const context = await spo.ensureFormDigest(spoAdminUrl as string, logger, undefined, verbose);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': (context as FormDigestInfo).FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: any = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      if (response.ErrorInfo.ErrorTypeName === 'Microsoft.SharePoint.Client.UnknownError') {
        return false;
      }

      throw response.ErrorInfo.ErrorMessage;
    }

    const site: DeletedSiteProperties = json[json.length - 1];

    return site.Status === 'Recycled';
  },

  /**
  * Deletes a site from the recycle bin
  * @param url The url of the site
  * @param logger the Logger object
  * @param verbose set if verbose logging should be logged
  * @param wait set to wait until finished
  */
  async deleteSiteFromTheRecycleBin(url: string, logger: Logger, verbose: boolean, wait?: boolean): Promise<void> {
    const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
    const context = await spo.ensureFormDigest(spoAdminUrl as string, logger, undefined, verbose);

    if (verbose) {
      await logger.logToStderr(`Deleting site ${url} from the recycle bin...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': context.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const response: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }

    const operation: SpoOperation = json[json.length - 1];
    const isComplete: boolean = operation.IsComplete;

    if (!wait || isComplete) {
      return;
    }

    await setTimeout(operation.PollingInterval);
    await spo.waitUntilFinished({
      operationId: JSON.stringify(operation._ObjectIdentity_),
      siteUrl: spoAdminUrl,
      logger,
      currentContext: context,
      verbose: verbose,
      debug: verbose
    });
  },

  /**
   * Updates a site with the given properties
   * @param url The url of the site
   * @param logger The logger object
   * @param verbose Set for verbose logging
   * @param title The new title
   * @param classification The classification to be updated
   * @param disableFlows If flows should be disabled or not
   * @param isPublic If site should be public or private
   * @param owners The owners to be updated
   * @param shareByEmailEnabled If share by e-mail should be enabled or not
   * @param siteDesignId The site design to be updated
   * @param sharingCapability The sharing capability to be updated
   */
  async updateSite(url: string, logger: Logger, verbose: boolean, title?: string, classification?: string, disableFlows?: boolean, isPublic?: boolean, owners?: string, shareByEmailEnabled?: boolean, siteDesignId?: string, sharingCapability?: string): Promise<void> {
    const tenantId = await spo.getTenantId(logger, verbose);
    const spoAdminUrl = await spo.getSpoAdminUrl(logger, verbose);
    let context = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, verbose);

    if (verbose) {
      await logger.logToStderr('Loading site IDs...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${url}/_api/site?$select=GroupId,Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const siteInfo = await request.get<{ GroupId: string; Id: string }>(requestOptions);
    const groupId = siteInfo.GroupId;
    const siteId = siteInfo.Id;
    const isGroupConnectedSite = groupId !== '00000000-0000-0000-0000-000000000000';

    if (verbose) {
      await logger.logToStderr(`Retrieved site IDs. siteId: ${siteId}, groupId: ${groupId}`);
    }

    if (isGroupConnectedSite) {
      if (verbose) {
        await logger.logToStderr(`Site attached to group ${groupId}`);
      }

      if (typeof title !== 'undefined' &&
        typeof isPublic !== 'undefined' &&
        typeof owners !== 'undefined') {


        const promises: Promise<void>[] = [];

        if (typeof title !== 'undefined') {
          const requestOptions: CliRequestOptions = {
            url: `${spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
            headers: {
              accept: 'application/json;odata=nometadata',
              'content-type': 'application/json;charset=utf-8',
              'X-RequestDigest': context.FormDigestValue
            },
            data: {
              groupId: groupId,
              siteId: siteId,
              displayName: title
            },
            responseType: 'json'
          };

          promises.push(request.post(requestOptions));
        }

        if (typeof isPublic !== 'undefined') {
          promises.push(entraGroup.setGroup(groupId as string, (isPublic === false), logger, verbose));
        }
        if (typeof owners !== 'undefined') {
          promises.push(spo.setGroupifiedSiteOwners(spoAdminUrl, groupId, owners, logger, verbose));
        }

        await Promise.all(promises);
      }
    }
    else {
      if (verbose) {
        await logger.logToStderr('Site is not group connected');
      }

      if (typeof isPublic !== 'undefined') {
        throw `The isPublic option can't be set on a site that is not groupified`;
      }

      if (owners) {
        await Promise.all(owners!.split(',').map(async (o) => {
          await spo.setSiteAdmin(spoAdminUrl, context, url, o.trim());
        }));
      }
    }

    context = await spo.ensureFormDigest(spoAdminUrl, logger, context, verbose);

    if (verbose) {
      await logger.logToStderr(`Updating site ${url} properties...`);
    }

    let updatedProperties: boolean = false;

    if (!isGroupConnectedSite) {
      if (title !== undefined) {
        updatedProperties = true;
      }
    }
    else {
      if (classification !== undefined || disableFlows !== undefined || shareByEmailEnabled !== undefined || sharingCapability !== undefined) {
        updatedProperties = true;
      }
    }

    if (updatedProperties) {
      let propertyId: number = 27;
      const payload: string[] = [];

      if (!isGroupConnectedSite) {
        if (title) {
          payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Title"><Parameter Type="String">${formatting.escapeXml(title)}</Parameter></SetProperty>`);
        }
      }
      if (typeof classification === 'string') {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${formatting.escapeXml(classification)}</Parameter></SetProperty>`);
      }
      if (typeof disableFlows !== 'undefined') {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">${disableFlows}</Parameter></SetProperty>`);
      }
      if (typeof shareByEmailEnabled !== 'undefined') {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">${shareByEmailEnabled}</Parameter></SetProperty>`);
      }
      if (sharingCapability) {
        const sharingCapabilityOption: SharingCapabilities = SharingCapabilities[(sharingCapability as keyof typeof SharingCapabilities)];
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">${sharingCapabilityOption}</Parameter></SetProperty>`);
      }

      const pos: number = (tenantId as string).indexOf('|') + 1;

      const requestOptionsUpdateProperties: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${(tenantId as string).substr(pos, (tenantId as string).indexOf('&') - pos)}&#xA;SiteProperties&#xA;${formatting.encodeQueryParameter(url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`

      };

      const res = await request.post<string>(requestOptionsUpdateProperties);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const operation: SpoOperation = json[json.length - 1];
        const isComplete: boolean = operation.IsComplete;
        if (!isComplete) {
          await setTimeout(operation.PollingInterval);
          await spo.waitUntilFinished({
            operationId: JSON.stringify(operation._ObjectIdentity_),
            siteUrl: spoAdminUrl,
            logger,
            currentContext: context,
            verbose: verbose,
            debug: verbose
          });
        }

      }
    }
    if (siteDesignId) {
      await spo.applySiteDesign(siteDesignId, url, logger, verbose);
    }
  },

  /**
   * Updates the groupified site owners
   * @param spoAdminUrl The SharePoint admin url
   * @param groupId The ID of the group
   * @param owners The owners to be updated
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async setGroupifiedSiteOwners(spoAdminUrl: string, groupId: string, owners: string, logger: Logger, verbose: boolean): Promise<void> {
    const splittedOwners: string[] = owners.split(',').map(o => o.trim());

    if (verbose) {
      await logger.logToStderr('Retrieving user information to set group owners...');
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/users?$filter=${splittedOwners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (res.value.length === 0) {
      return;
    }

    await Promise.all(res.value.map(user => {
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SP.Directory.DirectorySession/Group('${groupId}')/Owners/Add(objectId='${user.id}', principalName='')`,
        headers: {
          'content-type': 'application/json;odata=verbose'
        }
      };

      return request.post(requestOptions);
    }));
  },

  /**
   * Updates the site admin
   * @param spoAdminUrl The SharePoint admin url
   * @param context The FormDigestInfo
   * @param siteUrl The url of the site where the owners has to be updated
   * @param principal The principal of the admin
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async setSiteAdmin(spoAdminUrl: string, context: FormDigestInfo, siteUrl: string, principal: string, logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Updating site admin ${principal} on ${siteUrl}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': context.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${formatting.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${formatting.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post<string>(requestOptions);

    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];
    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }
    else {
      return;
    }
  },

  /**
   * Applies a site design
   * @param id The ID of the site design (GUID)
   * @param webUrl The url of the site where the design should be applied
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async applySiteDesign(id: string, webUrl: string, logger: Logger, verbose: boolean): Promise<void> {
    if (verbose) {
      await logger.logToStderr(`Applying site design ${id}...`);
    }

    const spoUrl: string = await spo.getSpoUrl(logger, verbose);

    const requestOptions: CliRequestOptions = {
      url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ApplySiteDesign`,
      headers: {
        'content-type': 'application/json;charset=utf-8',
        accept: 'application/json;odata=nometadata'
      },
      data: {
        siteDesignId: id,
        webUrl: webUrl
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  },

  /**
   * Gets the web properties for a given url
   * @param url The url of the web
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async getWeb(url: string, logger?: Logger, verbose?: boolean): Promise<WebProperties> {
    if (verbose && logger) {
      await logger.logToStderr(`Getting the web properties for ${url}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${url}/_api/web`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const webProperties: WebProperties = await request.get<WebProperties>(requestOptions);

    return webProperties;
  },

  /**
   * Applies the retention label to the items in the given list.
   * @param webUrl The url of the web
   * @param name The name of the label
   * @param listAbsoluteUrl The absolute Url to the list
   * @param itemIds The list item Ids to apply the label to. (A maximum 100 is allowed)
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async applyRetentionLabelToListItems(webUrl: string, name: string, listAbsoluteUrl: string, itemIds: number[], logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Applying retention label '${name}' to item(s) in list '${listAbsoluteUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        listUrl: listAbsoluteUrl,
        complianceTagValue: name,
        itemIds: itemIds
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  },

  /**
  * Gets a file as list item by url
  * @param absoluteListUrl The absolute url to the list
  * @param url The url of the file
  * @param logger The logger object
  * @param verbose If in verbose mode
  * @returns The list item object
  */
  async getFileAsListItemByUrl(absoluteListUrl: string, url: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      await logger.logToStderr(`Getting the file properties with url ${url}`);
    }

    const serverRelativePath = urlUtil.getServerRelativePath(absoluteListUrl, url);
    const requestUrl = `${absoluteListUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?$expand=ListItemAllFields&@f='${formatting.encodeQueryParameter(serverRelativePath)}'`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const file = await request.get<FileProperties>(requestOptions);

    return file.ListItemAllFields;
  },

  /**
  * Updates a list item with system update
  * @param absoluteListUrl The absolute base URL without query parameters, pointing to the specific list where the item resides. This URL should represent the list.
  * @param itemId The id of the list item
  * @param properties An object of the properties that should be updated
  * @param contentTypeName The name of the content type to update
  * @param logger The logger object
  * @param verbose If in verbose mode
  * @returns The updated list item object
  */
  async systemUpdateListItem(absoluteListUrl: string, itemId: string, logger: Logger, verbose: boolean, properties?: object, contentTypeName?: string): Promise<ListItemInstance> {
    if (!properties && !contentTypeName) {
      // Neither properties nor contentTypeName provided, no need to proceed
      throw 'Either properties or contentTypeName must be provided for systemUpdateListItem.';
    }

    const parsedUrl = new URL(absoluteListUrl);
    const serverRelativeSiteMatch = absoluteListUrl.match(new RegExp('/sites/[^/]+'));
    const webUrl = `${parsedUrl.protocol}//${parsedUrl.host}${serverRelativeSiteMatch ?? ''}`;

    if (verbose && logger) {
      await logger.logToStderr(`Getting list id...`);
    }

    const listRequestOptions: CliRequestOptions = {
      url: `${absoluteListUrl}?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const list = await request.get<{ Id: string; }>(listRequestOptions);

    const listId = list.Id;

    if (verbose && logger) {
      await logger.logToStderr(`Getting request digest for systemUpdate request`);
    }

    const res = await spo.getRequestDigest(webUrl);

    const formDigestValue = res.FormDigestValue;
    const objectIdentity: string = await spo.requestObjectIdentity(webUrl, logger, verbose);

    let index = 0;

    const requestBodyOptions: string[] = properties ? Object.keys(properties).map(key => `
    <Method Name="ParseAndSetFieldValue" Id="${++index}" ObjectPathId="147">
      <Parameters>
        <Parameter Type="String">${key}</Parameter>
        <Parameter Type="String">${(<any>properties)[key].toString()}</Parameter>
      </Parameters>
    </Method>`) : [];

    const additionalContentType: string = contentTypeName ? `
    <Method Name="ParseAndSetFieldValue" Id="${++index}" ObjectPathId="147">
      <Parameters>
        <Parameter Type="String">ContentType</Parameter>
        <Parameter Type="String">${contentTypeName}</Parameter>
      </Parameters>
    </Method>` : '';

    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
      <Actions>
        ${requestBodyOptions.join('')}${additionalContentType}
        <Method Name="SystemUpdate" Id="${++index}" ObjectPathId="147" />
      </Actions>
      <ObjectPaths>
        <Identity Id="147" Name="${objectIdentity}:list:${listId}:item:${itemId},1" />
      </ObjectPaths>
    </Request>`;

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'Content-Type': 'text/xml',
        'X-RequestDigest': formDigestValue
      },
      data: requestBody
    };

    const response: string = await request.post(requestOptions);

    if (response.indexOf("ErrorMessage") > -1) {
      throw `Error occurred in systemUpdate operation - ${response}`;
    }

    const id = Number(itemId);

    const requestOptionsItems: CliRequestOptions = {
      url: `${absoluteListUrl}/items(${id})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const itemsResponse = await request.get<ListItemInstance>(requestOptionsItems);
    return (itemsResponse);
  },

  /**
   * Removes the retention label from the items in the given list.
   * @param webUrl The url of the web
   * @param listAbsoluteUrl The absolute Url to the list
   * @param itemIds The list item Ids to clear the label from (A maximum 100 is allowed)
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async removeRetentionLabelFromListItems(webUrl: string, listAbsoluteUrl: string, itemIds: number[], logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Removing retention label from item(s) in list '${listAbsoluteUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        listUrl: listAbsoluteUrl,
        complianceTagValue: '',
        itemIds: itemIds
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  },

  /**
   * Applies a default retention label to a list or library.
   * @param webUrl The url of the web
   * @param name The name of the label
   * @param listAbsoluteUrl The absolute Url to the list
   * @param syncToItems If the label needs to be synced to existing items/files.
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async applyDefaultRetentionLabelToList(webUrl: string, name: string, listAbsoluteUrl: string, syncToItems?: boolean, logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Applying default retention label '${name}' to the list '${listAbsoluteUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        listUrl: listAbsoluteUrl,
        complianceTagValue: name,
        blockDelete: false,
        blockEdit: false,
        syncToItems: syncToItems || false
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  },

  /**
   * Removes the default retention label from a list or library.
   * @param webUrl The url of the web
   * @param listAbsoluteUrl The absolute Url to the list
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async removeDefaultRetentionLabelFromList(webUrl: string, listAbsoluteUrl: string, logger?: Logger, verbose?: boolean): Promise<void> {
    if (verbose && logger) {
      await logger.logToStderr(`Removing the default retention label from the list '${listAbsoluteUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        listUrl: listAbsoluteUrl,
        complianceTagValue: '',
        blockDelete: false,
        blockEdit: false,
        syncToItems: false
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  },


  /**
   * Retrieves the site ID for a given web URL.
   * @param webUrl The web URL for which to retrieve the site ID.
   * @param logger The logger object.
   * @param verbose Set for verbose logging
   * @returns A promise that resolves to the site ID.
   */
  async getSiteId(webUrl: string, logger?: Logger, verbose?: boolean): Promise<string> {
    if (verbose && logger) {
      await logger.logToStderr(`Getting site id for URL: ${webUrl}...`);
    }

    const url: URL = new URL(webUrl);
    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/sites/${formatting.encodeQueryParameter(url.host)}:${url.pathname}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const site: Site = await request.get<Site>(requestOptions);

    return site.id as string;
  },

  /**
   * Retrieves the server-relative URL of a folder.
   * @param webUrl Web URL
   * @param folderUrl Folder URL
   * @param folderId Folder ID
   * @param logger The logger object
   * @param verbose Set for verbose logging
   * @returns The server-relative URL of the folder
   */
  async getFolderServerRelativeUrl(webUrl: string, folderUrl?: string, folderId?: string, logger?: Logger, verbose?: boolean): Promise<string> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving server-relative URL for folder ${folderUrl ? `URL: ${folderUrl}` : `ID: ${folderId}`}`);
    }

    let requestUrl: string = `${webUrl}/_api/web/`;

    if (folderUrl) {
      const folderServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, folderUrl);
      requestUrl += `GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')`;
    }
    else {
      requestUrl += `GetFolderById('${folderId}')`;
    }

    requestUrl += '?$select=ServerRelativeUrl';

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<{ ServerRelativeUrl: string }>(requestOptions);
    return res.ServerRelativeUrl;
  },

  /**
   * Retrieves the ObjectIdentity from a SharePoint site
   * @param webUrl web url
   * @param logger The logger object
   * @param verbose If in verbose mode
   * @return The ObjectIdentity as string
   */
  async requestObjectIdentity(webUrl: string, logger: Logger, verbose: boolean): Promise<string> {
    const res = await spo.getRequestDigest(webUrl);
    const formDigestValue = res.FormDigestValue;

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    const response = await request.post<string>(requestOptions);
    if (verbose) {
      await logger.logToStderr('Attempt to get _ObjectIdentity_ key values');
    }

    const json: ClientSvcResponse = JSON.parse(response);

    const contents: ClientSvcResponseContents = json.find(x => { return x.ErrorInfo; });
    if (contents && contents.ErrorInfo) {
      throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
    }

    const identityObject = json.find(x => { return x._ObjectIdentity_; });
    if (identityObject) {
      return identityObject._ObjectIdentity_;
    }

    throw 'Cannot proceed. _ObjectIdentity_ not found'; // this is not supposed to happen
  },

  /**
  * Updates a list item without system update
  * @param absoluteListUrl The absolute base URL without query parameters, pointing to the specific list where the item resides. This URL should represent the list.
  * @param itemId The id of the list item
  * @param properties An object of the properties that should be updated
  * @param contentTypeName The name of the content type to update
  * @returns The updated listitem object
  */
  async updateListItem(absoluteListUrl: string, itemId: string, properties?: object, contentTypeName?: string): Promise<ListItemInstance> {
    const requestBodyOptions: any[] = [
      ...(properties
        ? Object.keys(properties).map((key: string) => ({
          FieldName: key,
          FieldValue: (<any>properties)[key].toString()
        }))
        : [])
    ];

    const requestBody: {
      formValues: FormValues[]
    } = {
      formValues: requestBodyOptions
    };

    contentTypeName && requestBody.formValues.push({
      FieldName: 'ContentType',
      FieldValue: contentTypeName
    });

    const requestOptions: CliRequestOptions = {
      url: `${absoluteListUrl}/items(${itemId})/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    const response = await request.post<{ value: ListItemFieldValueResult[] }>(requestOptions);

    // Response is from /ValidateUpdateListItem POST call, perform get on updated item to get all field values
    const fieldValues: ListItemFieldValueResult[] = response.value;
    if (fieldValues.some(f => f.HasException)) {
      throw `Updating the items has failed with the following errors: ${os.EOL}${fieldValues.filter(f => f.HasException).map(f => { return `- ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`;
    }

    const requestOptionsItems: CliRequestOptions = {
      url: `${absoluteListUrl}/items(${itemId})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const itemsResponse = await request.get<ListItemInstance>(requestOptionsItems);
    return (itemsResponse);
  },

  /**
  * Retrieves the file by id.
  * Returns a FileProperties object
  * @param webUrl Web url
  * @param id the id of the file
  * @param logger the Logger object
  * @param verbose set for verbose logging 
  */
  async getFileById(webUrl: string, id: string, logger?: Logger, verbose?: boolean): Promise<FileProperties> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving the file with id ${id}`);
    }
    const requestUrl = `${webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(id)}')`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },

      responseType: 'json'
    };

    const file: FileProperties = await request.get<FileProperties>(requestOptions);

    return file;
  },

  /**
   * Gets the primary owner login from a site as admin.
   * @param adminUrl The SharePoint admin URL
   * @param siteId The site ID
   * @param logger The logger object
   * @param verbose If in verbose mode
   * @returns Owner login name
   */
  async getPrimaryAdminLoginNameAsAdmin(adminUrl: string, siteId: string, logger: Logger, verbose: boolean): Promise<string> {
    if (verbose) {
      await logger.logToStderr('Getting the primary admin login name...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ OwnerLoginName: string }>(requestOptions);
    return response.OwnerLoginName;
  },

  /**
   * Gets the primary owner login from a site.
   * @param siteUrl The site URL
   * @param logger The logger object
   * @param verbose If in verbose mode
   * @returns Owner login name
   */
  async getPrimaryOwnerLoginFromSite(siteUrl: string, logger: Logger, verbose: boolean): Promise<string> {
    if (verbose) {
      await logger.logToStderr('Getting the primary admin login name...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/site/owner`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const responseContent = await request.get<{ LoginName: string }>(requestOptions);
    return responseContent?.LoginName;
  }
};