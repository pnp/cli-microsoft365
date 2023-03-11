import * as url from 'url';
import { urlUtil } from "./urlUtil";
import { validation } from "./validation";
import auth from '../Auth';
import { Logger } from "../cli/Logger";
import config from "../config";
import { BasePermissions } from '../m365/spo/base-permissions';
import request, { CliRequestOptions } from "../request";
import { formatting } from './formatting';
import { CustomAction } from '../m365/spo/commands/customaction/customaction';
import { odata } from './odata';
import { ListItemRetentionLabel } from '../m365/spo/commands/listitem/ListItemRetentionLabel';
import { SiteRetentionLabel } from '../m365/spo/commands/listitem/SiteRetentionLabel';

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

export interface SearchResponse {
  PrimaryQueryResult: {
    RelevantResults: {
      RowCount: number;
      Table: {
        Rows: {
          Cells: {
            Key: string;
            Value: string;
            ValueType: string;
          }[];
        }[];
      };
    }
  }
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

export const spo = {
  getRequestDigest(siteUrl: string): Promise<FormDigestInfo> {
    const requestOptions: any = {
      url: `${siteUrl}/_api/contextinfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  },

  ensureFormDigest(siteUrl: string, logger: Logger, context: FormDigestInfo | undefined, debug: boolean): Promise<FormDigestInfo> {
    return new Promise<FormDigestInfo>((resolve: (context: FormDigestInfo) => void, reject: (error: any) => void): void => {
      if (validation.isValidFormDigest(context)) {
        if (debug) {
          logger.logToStderr('Existing form digest still valid');
        }

        resolve(context as FormDigestInfo);
        return;
      }

      spo
        .getRequestDigest(siteUrl)
        .then((res: FormDigestInfo): void => {
          const now: Date = new Date();
          now.setSeconds(now.getSeconds() + res.FormDigestTimeoutSeconds - 5);
          context = {
            FormDigestValue: res.FormDigestValue,
            FormDigestTimeoutSeconds: res.FormDigestTimeoutSeconds,
            FormDigestExpiresAt: now,
            WebFullUrl: res.WebFullUrl
          };

          resolve(context);
        }, (error: any): void => {
          reject(error);
        });
    });
  },

  waitUntilFinished({ operationId, siteUrl, resolve, reject, logger, currentContext, debug, verbose }: { operationId: string, siteUrl: string, resolve: () => void, reject: (error: any) => void, logger: Logger, currentContext: FormDigestInfo, debug: boolean, verbose: boolean }): void {
    spo
      .ensureFormDigest(siteUrl, logger, currentContext, debug)
      .then((res: FormDigestInfo): Promise<string> => {
        currentContext = res;

        if (debug) {
          logger.logToStderr(`Checking if operation ${operationId} completed...`);
        }

        const requestOptions: any = {
          url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': currentContext.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="${operationId.replace(/\\n/g, '&#xA;').replace(/"/g, '')}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          reject(response.ErrorInfo.ErrorMessage);
        }
        else {
          const operation: SpoOperation = json[json.length - 1];
          const isComplete: boolean = operation.IsComplete;
          if (isComplete) {
            if (!debug && verbose) {
              process.stdout.write('\n');
            }

            resolve();
            return;
          }

          setTimeout(() => {
            spo.waitUntilFinished({
              operationId: JSON.stringify(operation._ObjectIdentity_),
              siteUrl,
              resolve,
              reject,
              logger,
              currentContext,
              debug,
              verbose
            });
          }, operation.PollingInterval);
        }
      });
  },

  waitUntilCopyJobFinished({ copyJobInfo, siteUrl, pollingInterval, resolve, reject, logger, debug, verbose }: { copyJobInfo: any, siteUrl: string, pollingInterval: number, resolve: () => void, reject: (error: any) => void, logger: Logger, debug: boolean, verbose: boolean }): void {
    const requestUrl: string = `${siteUrl}/_api/site/GetCopyJobProgress`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: { "copyJobInfo": copyJobInfo },
      responseType: 'json'
    };

    request
      .post<{ JobState?: number, Logs: string[] }>(requestOptions)
      .then((resp: { JobState?: number, Logs: string[] }): void => {
        if (debug) {
          logger.logToStderr('getCopyJobProgress response...');
          logger.logToStderr(resp);
        }

        for (const item of resp.Logs) {
          const log: { Event: string; Message: string } = JSON.parse(item);

          // reject if progress error
          if (log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // two possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        if (resp.JobState === 0) {
          // job done
          if (verbose) {
            process.stdout.write('\n');
          }

          resolve();
        }
        else {
          setTimeout(() => {
            spo.waitUntilCopyJobFinished({ copyJobInfo, siteUrl, pollingInterval, resolve, reject, logger, debug, verbose });
          }, pollingInterval);
        }
      });
  },

  getSpoUrl(logger: Logger, debug: boolean): Promise<string> {
    if (auth.service.spoUrl) {
      if (debug) {
        logger.logToStderr(`SPO URL previously retrieved ${auth.service.spoUrl}. Returning...`);
      }

      return Promise.resolve(auth.service.spoUrl);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`No SPO URL available. Retrieving from MS Graph...`);
      }

      const requestOptions: any = {
        url: `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .get<{ webUrl: string }>(requestOptions)
        .then((res: { webUrl: string }): Promise<void> => {
          auth.service.spoUrl = res.webUrl;
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          resolve(auth.service.spoUrl as string);
        }, (err: any): void => {
          if (auth.service.spoUrl) {
            resolve(auth.service.spoUrl);
          }
          else {
            reject(err);
          }
        });
    });
  },

  getSpoAdminUrl(logger: Logger, debug: boolean): Promise<string> {
    return new Promise<string>((resolve: (spoAdminUrl: string) => void, reject: (error: any) => void): void => {
      spo
        .getSpoUrl(logger, debug)
        .then((spoUrl: string): void => {
          resolve(spoUrl.replace(/(https:\/\/)([^\.]+)(.*)/, '$1$2-admin$3'));
        }, (error: any): void => {
          reject(error);
        });
    });
  },

  getTenantId(logger: Logger, debug: boolean): Promise<string> {
    if (auth.service.tenantId) {
      if (debug) {
        logger.logToStderr(`SPO Tenant ID previously retrieved ${auth.service.tenantId}. Returning...`);
      }

      return Promise.resolve(auth.service.tenantId);
    }

    return new Promise<string>((resolve: (spoUrl: string) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`No SPO Tenant ID available. Retrieving...`);
      }

      let spoAdminUrl: string = '';

      spo
        .getSpoAdminUrl(logger, debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return spo.getRequestDigest(spoAdminUrl);
        })
        .then((contextInfo: ContextInfo): Promise<string> => {
          const tenantInfoRequestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': contextInfo.FormDigestValue,
              accept: 'application/json;odata=nometadata'
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
          };

          return request.post(tenantInfoRequestOptions);
        })
        .then((res: string): Promise<void> => {
          const json: string[] = JSON.parse(res);
          auth.service.tenantId = (json[json.length - 1] as any)._ObjectIdentity_.replace('\n', '&#xA;');
          return auth.storeConnectionInfo();
        })
        .then((): void => {
          resolve(auth.service.tenantId as string);
        }, (err: any): void => {
          if (auth.service.tenantId) {
            resolve(auth.service.tenantId);
          }
          else {
            reject(err);
          }
        });
    });
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
  ensureFolder(webFullUrl: string, folderToEnsure: string, logger: Logger, debug: boolean): Promise<void> {
    const webUrl = url.parse(webFullUrl);
    if (!webUrl.protocol || !webUrl.hostname) {
      return Promise.reject('webFullUrl is not a valid URL');
    }

    if (!folderToEnsure) {
      return Promise.reject('folderToEnsure cannot be empty');
    }

    // remove last '/' of webFullUrl if exists
    const webFullUrlLastCharPos: number = webFullUrl.length - 1;

    if (webFullUrl.length > 1 &&
      webFullUrl[webFullUrlLastCharPos] === '/') {
      webFullUrl = webFullUrl.substring(0, webFullUrlLastCharPos);
    }

    folderToEnsure = urlUtil.getWebRelativePath(webFullUrl, folderToEnsure);

    if (debug) {
      logger.log(`folderToEnsure`);
      logger.log(folderToEnsure);
      logger.log('');
    }

    let nextFolder: string = '';
    let prevFolder: string = '';
    let folderIndex: number = 0;

    // build array of folders e.g. ["Shared%20Documents","22","54","55"]
    const folders: string[] = folderToEnsure.substring(1).split('/');

    if (debug) {
      logger.log('folders to process');
      logger.log(JSON.stringify(folders));
      logger.log('');
    }

    // recursive function
    const checkOrAddFolder = (resolve: () => void, reject: (error: any) => void): void => {
      if (folderIndex === folders.length) {
        if (debug) {
          logger.log(`All sub-folders exist`);
        }

        return resolve();
      }

      // append the next sub-folder to the folder path and check if it exists
      prevFolder = nextFolder;
      nextFolder += `/${folders[folderIndex]}`;
      const folderServerRelativeUrl = urlUtil.getServerRelativePath(webFullUrl, nextFolder);

      const requestOptions: any = {
        url: `${webFullUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderServerRelativeUrl)}')`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        }
      };

      request
        .get(requestOptions)
        .then(() => {
          folderIndex++;
          checkOrAddFolder(resolve, reject);
        })
        .catch(() => {
          const prevFolderServerRelativeUrl: string = urlUtil.getServerRelativePath(webFullUrl, prevFolder);
          const requestOptions: any = {
            url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27${formatting.encodeQueryParameter(prevFolderServerRelativeUrl)}%27&@a2=%27${formatting.encodeQueryParameter(folders[folderIndex])}%27`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.post(requestOptions)
            .then(() => {
              folderIndex++;
              checkOrAddFolder(resolve, reject);
            })
            .catch((err: any) => {
              if (debug) {
                logger.log(`Could not create sub-folder ${folderServerRelativeUrl}`);
              }

              reject(err);
            });
        });
    };
    return new Promise<void>(checkOrAddFolder);
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
  getCurrentWebIdentity(webUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
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
  },

  /**
   * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
   * @param webObjectIdentity ObjectIdentity. Has format _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * @param webUrl web url
   * @param siteAccessToken site access token
   * @param formDigestValue formDigestValue
   */
  getEffectiveBasePermissions(webObjectIdentity: string, webUrl: string, formDigestValue: string, logger: Logger, debug: boolean): Promise<BasePermissions> {
    const basePermissionsResult: BasePermissions = new BasePermissions();

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    return new Promise<BasePermissions>((resolve: (permissions: BasePermissions) => void, reject: (error: any) => void): void => {
      request.post<string>(requestOptions).then((res: string): void => {
        if (debug) {
          logger.log('Attempt to get the web EffectiveBasePermissions');
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
      }, (err: any): void => { reject(err); });
    });
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

  getFolderIdentity(webObjectIdentity: string, webUrl: string, siteRelativeUrl: string, formDigestValue: string): Promise<IdentityResponse> {
    const serverRelativePath: string = urlUtil.getServerRelativePath(webUrl, siteRelativeUrl);

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativePath}</Parameter></Parameters></Method><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
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
      }, (err: any): void => { reject(err); });
    });
  },

  /**
   * Retrieves the SiteId, VroomItemId and VroomDriveId from a specific file.
   * @param webUrl Web url
   * @param fileId GUID ID of the file
   * @param fileUrl Decoded URL of the file
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
      const requestOptions: any = {
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

    const requestOptions: any = {
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
   * Retrieves the Azure AD ID from a SP user.
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

  async getWebRetentionLabelInformationByName(webUrl: string, name: string): Promise<ListItemRetentionLabel> {

    const requestUrl: string = `${webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`;

    const labels: SiteRetentionLabel[] = await odata.getAllItems(requestUrl);

    const label = labels.find(l => l.TagName === name);

    if (label === undefined) {
      throw new Error(`The specified retention label does not exist`);
    }

    return {
      complianceTag: label.TagName,
      isTagPolicyHold: label.BlockDelete,
      isTagPolicyRecord: label.BlockEdit,
      isEventBasedTag: label.IsEventTag,
      isTagSuperLock: label.SuperLock,
      isUnlockedAsDefault: label.UnlockedAsDefault
    } as ListItemRetentionLabel;
  }
};
