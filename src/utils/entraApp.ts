import { Application } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";

const graphResource = 'https://graph.microsoft.com';

export interface AppInfo extends Application {
  appId: string;
  id: string;
  tenantId: string;
  secrets?: {
    displayName: string;
    value: string;
  }[];
  requiredResourceAccess: RequiredResourceAccess[];
}

export interface AppPermissions {
  resourceId: string;
  resourceAccess: ResourceAccess[];
  scope: string[];
}

export interface ResourceAccess {
  id: string;
  type: string;
}

export interface ServicePrincipalInfo {
  appId: string;
  appRoles: { id: string; value: string; }[];
  id: string;
  oauth2PermissionScopes: { id: string; value: string; }[];
  servicePrincipalNames: string[];
}

export interface RequiredResourceAccess {
  resourceAppId: string;
  resourceAccess: ResourceAccess[];
}

export const entraApp = {

  async createEntraApp(applicationInfo: Application): Promise<AppInfo> {
    const createApplicationRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/applications`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request.post<AppInfo>(createApplicationRequestOptions);
  },

  async updateEntraApp(appId: string, applicationInfo: Application): Promise<AppInfo> {
    const updateAppRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/applications/${appId}`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: applicationInfo
    };
    return await request.patch<AppInfo>(updateAppRequestOptions);
  },

  async addRoleToServicePrincipal(objectId: string, resourceId: string, appRoleId: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/servicePrincipals/${objectId}/appRoleAssignments`,
      headers: {
        'Content-Type': 'application/json'
      },
      responseType: 'json',
      data: {
        appRoleId: appRoleId,
        principalId: objectId,
        resourceId: resourceId
      }
    };

    return await request.post(requestOptions);
  },

  async grantOAuth2Permission(appId: string, resourceId: string, scopeName: string): Promise<void> {
    const grantAdminConsentApplicationRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/oauth2PermissionGrants`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        clientId: appId,
        consentType: "AllPrincipals",
        principalId: null,
        resourceId: resourceId,
        scope: scopeName
      }
    };

    return await request.post(grantAdminConsentApplicationRequestOptions);
  },

  async createServicePrincipal(appId: string): Promise<ServicePrincipalInfo> {
    const requestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/servicePrincipals`,
      headers: {
        'content-type': 'application/json'
      },
      data: {
        appId: appId
      },
      responseType: 'json'
    };

    return await request.post<ServicePrincipalInfo>(requestOptions);
  }
};
