import * as url from 'url';
import { urlUtil, validation } from ".";
import auth from '../Auth';
import { Logger } from "../cli";
import config from "../config";
import { BasePermissions } from '../m365/spo/base-permissions';
import request from "../request";

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

  waitUntilFinished({ operationId, siteUrl, resolve, reject, logger, currentContext, dots, debug, verbose }: { operationId: string, siteUrl: string, resolve: () => void, reject: (error: any) => void, logger: Logger, currentContext: FormDigestInfo, dots?: string, debug: boolean, verbose: boolean }): void {
    spo
      .ensureFormDigest(siteUrl, logger, currentContext, debug)
      .then((res: FormDigestInfo): Promise<string> => {
        currentContext = res;

        if (debug) {
          logger.logToStderr(`Checking if operation ${operationId} completed...`);
        }

        if (!debug && verbose) {
          dots += '.';
          process.stdout.write(`\r${dots}`);
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
              dots,
              debug,
              verbose
            });
          }, operation.PollingInterval);
        }
      });
  },

  waitUntilCopyJobFinished({ copyJobInfo, siteUrl, pollingInterval, resolve, reject, logger, dots, debug, verbose }: { copyJobInfo: any, siteUrl: string, pollingInterval: number, resolve: () => void, reject: (error: any) => void, logger: Logger, dots?: string, debug: boolean, verbose: boolean }): void {
    const requestUrl: string = `${siteUrl}/_api/site/GetCopyJobProgress`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: { "copyJobInfo": copyJobInfo },
      responseType: 'json'
    };

    if (!debug && verbose) {
      dots += '.';
      process.stdout.write(`\r${dots}`);
    }

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
            spo.waitUntilCopyJobFinished({ copyJobInfo, siteUrl, pollingInterval, resolve, reject, logger, dots, debug, verbose });
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
        url: `${webFullUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')`,
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
            url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27${encodeURIComponent(prevFolderServerRelativeUrl)}%27&@a2=%27${encodeURIComponent(folders[folderIndex])}%27`,
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
  }
};