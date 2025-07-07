import { Application, RequiredResourceAccess, ResourceAccess } from '@microsoft/microsoft-graph-types';
import fs from 'fs';
import { Logger } from '../cli/Logger.js';
import request, { CliRequestOptions } from '../request.js';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import { optionsUtils } from './optionsUtils.js';

export interface AppInfo {
  appId: string;
  // objectId
  id: string;
  tenantId: string;
  secrets?: {
    displayName: string;
    value: string;
  }[];
  requiredResourceAccess: RequiredResourceAccess[];
}

export interface ServicePrincipalInfo {
  appId: string;
  appRoles: { id: string; value: string; }[];
  id: string;
  oauth2PermissionScopes: { id: string; value: string; }[];
  servicePrincipalNames: string[];
}

export interface AppCreationOptions {
  apisApplication?: string;
  apisDelegated?: string;
  implicitFlow: boolean;
  multitenant: boolean;
  name?: string;
  platform?: string;
  redirectUris?: string;
  certificateFile?: string;
  certificateBase64Encoded?: string;
  certificateDisplayName?: string;
  allowPublicClientFlows?: boolean;
  bundleId?: string;
  signatureHash?: string;
}

export interface AppPermissions {
  resourceId: string;
  resourceAccess: ResourceAccess[];
  scope: string[];
}

async function getCertificateBase64Encoded({ options, logger, debug }: {
  options: AppCreationOptions,
  logger: Logger,
  debug: boolean
}): Promise<string> {
  if (options.certificateBase64Encoded) {
    return options.certificateBase64Encoded;
  }

  if (debug) {
    await logger.logToStderr(`Reading existing ${options.certificateFile}...`);
  }

  try {
    return fs.readFileSync(options.certificateFile as string, { encoding: 'base64' });
  }
  catch (e) {
    throw new Error(`Error reading certificate file: ${e}. Please add the certificate using base64 option '--certificateBase64Encoded'.`);
  }
}

async function createServicePrincipal(appId: string): Promise<ServicePrincipalInfo> {
  const requestOptions: CliRequestOptions = {
    url: `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals`,
    headers: {
      'content-type': 'application/json'
    },
    data: {
      appId: appId
    },
    responseType: 'json'
  };

  return request.post<ServicePrincipalInfo>(requestOptions);
}

async function grantOAuth2Permission({ appId, resourceId, scopeName }: {
  appId: string,
  resourceId: string,
  scopeName: string
}): Promise<void> {
  const grantAdminConsentApplicationRequestOptions: CliRequestOptions = {
    url: `https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants`,
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

  return request.post(grantAdminConsentApplicationRequestOptions);
}

async function addRoleToServicePrincipal({ objectId, resourceId, appRoleId }: {
  objectId: string,
  resourceId: string,
  appRoleId: string
}): Promise<void> {
  const requestOptions: CliRequestOptions = {
    url: `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/${objectId}/appRoleAssignments`,
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

  return request.post(requestOptions);
}

async function getRequiredResourceAccessForApis({ servicePrincipals, apis, scopeType, logger, debug }: {
  servicePrincipals: ServicePrincipalInfo[],
  apis: string | undefined,
  scopeType: string,
  logger: Logger,
  debug: boolean
}): Promise<RequiredResourceAccess[]> {
  if (!apis) {
    return [];
  }

  const resolvedApis: RequiredResourceAccess[] = [];
  const requestedApis: string[] = apis!.split(',').map(a => a.trim());
  for (const api of requestedApis) {
    const pos: number = api.lastIndexOf('/');
    const permissionName: string = api.substring(pos + 1);
    const servicePrincipalName: string = api.substring(0, pos);
    if (debug) {
      await logger.logToStderr(`Resolving ${api}...`);
      await logger.logToStderr(`Permission name: ${permissionName}`);
      await logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
    }
    const servicePrincipal = servicePrincipals.find(sp => (
      sp.servicePrincipalNames.indexOf(servicePrincipalName) > -1 ||
      sp.servicePrincipalNames.indexOf(`${servicePrincipalName}/`) > -1));
    if (!servicePrincipal) {
      throw `Service principal ${servicePrincipalName} not found`;
    }

    const scopesOfType = scopeType === 'Scope' ? servicePrincipal.oauth2PermissionScopes : servicePrincipal.appRoles;
    const permission = scopesOfType.find(scope => scope.value === permissionName);
    if (!permission) {
      throw `Permission ${permissionName} for service principal ${servicePrincipalName} not found`;
    }

    let resolvedApi = resolvedApis.find(a => a.resourceAppId === servicePrincipal.appId);
    if (!resolvedApi) {
      resolvedApi = {
        resourceAppId: servicePrincipal.appId,
        resourceAccess: []
      };
      resolvedApis.push(resolvedApi);
    }

    const resourceAccessPermission = {
      id: permission.id,
      type: scopeType
    };

    resolvedApi.resourceAccess!.push(resourceAccessPermission);

    updateAppPermissions({
      spId: servicePrincipal.id,
      resourceAccessPermission,
      oAuth2PermissionValue: permission.value
    });
  }

  return resolvedApis;
}

function updateAppPermissions({ spId, resourceAccessPermission, oAuth2PermissionValue }: {
  spId: string,
  resourceAccessPermission: ResourceAccess,
  oAuth2PermissionValue?: string
}): void {
  // During API resolution, we store globally both app role assignments and oauth2permissions
  // So that we'll be able to parse them during the admin consent process
  let existingPermission = entraApp.appPermissions.find(oauth => oauth.resourceId === spId);
  if (!existingPermission) {
    existingPermission = {
      resourceId: spId,
      resourceAccess: [],
      scope: []
    };

    entraApp.appPermissions.push(existingPermission);
  }

  if (resourceAccessPermission.type === 'Scope' && oAuth2PermissionValue && !existingPermission.scope.find(scp => scp === oAuth2PermissionValue)) {
    existingPermission.scope.push(oAuth2PermissionValue);
  }

  if (!existingPermission.resourceAccess.find(res => res.id === resourceAccessPermission.id)) {
    existingPermission.resourceAccess.push(resourceAccessPermission);
  }
}

export const entraApp = {
  appPermissions: [] as AppPermissions[],
  createAppRegistration: async ({ options, apis, logger, verbose, debug, unknownOptions }: {
    options: AppCreationOptions,
    unknownOptions: any,
    apis: RequiredResourceAccess[],
    logger: Logger,
    verbose: boolean,
    debug: boolean
  }): Promise<AppInfo> => {
    const applicationInfo: any = {
      displayName: options.name,
      signInAudience: options.multitenant ? 'AzureADMultipleOrgs' : 'AzureADMyOrg'
    };

    if (apis.length > 0) {
      applicationInfo.requiredResourceAccess = apis;
    }

    if (options.redirectUris) {
      applicationInfo[options.platform!] = {
        redirectUris: options.redirectUris.split(',').map(u => u.trim())
      };
    }

    if (options.platform === 'android') {
      applicationInfo['publicClient'] = {
        redirectUris: [
          `msauth://${options.bundleId}/${formatting.encodeQueryParameter(options.signatureHash!)}`
        ]
      };
    }

    if (options.platform === 'apple') {
      applicationInfo['publicClient'] = {
        redirectUris: [
          `msauth://code/msauth.${options.bundleId}%3A%2F%2Fauth`,
          `msauth.${options.bundleId}://auth`
        ]
      };
    }

    if (options.implicitFlow) {
      if (!applicationInfo.web) {
        applicationInfo.web = {};
      }
      applicationInfo.web.implicitGrantSettings = {
        enableAccessTokenIssuance: true,
        enableIdTokenIssuance: true
      };
    }

    if (options.certificateFile || options.certificateBase64Encoded) {
      const certificateBase64Encoded = await getCertificateBase64Encoded({ options, logger, debug });

      const newKeyCredential = {
        type: 'AsymmetricX509Cert',
        usage: 'Verify',
        displayName: options.certificateDisplayName,
        key: certificateBase64Encoded
      } as any;

      applicationInfo.keyCredentials = [newKeyCredential];
    }

    if (options.allowPublicClientFlows) {
      applicationInfo.isFallbackPublicClient = true;
    }

    optionsUtils.addUnknownOptionsToPayload(applicationInfo, unknownOptions);

    if (verbose) {
      await logger.logToStderr(`Creating Microsoft Entra app registration...`);
    }

    const createApplicationRequestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/myorganization/applications`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request.post<AppInfo>(createApplicationRequestOptions);
  },
  grantAdminConsent: async ({ appInfo, appPermissions, adminConsent, logger, debug }: {
    appInfo: AppInfo,
    adminConsent: boolean | undefined,
    appPermissions: AppPermissions[],
    logger: Logger,
    debug: boolean
  }): Promise<AppInfo> => {
    if (!adminConsent || appPermissions.length === 0) {
      return appInfo;
    }

    const sp = await createServicePrincipal(appInfo.appId);
    if (debug) {
      await logger.logToStderr("Service principal created, returned object id: " + sp.id);
    }

    const tasks: Promise<void>[] = [];

    appPermissions.forEach(async (permission) => {
      if (permission.scope.length > 0) {
        tasks.push(grantOAuth2Permission({
          appId: sp.id,
          resourceId: permission.resourceId,
          scopeName: permission.scope.join(' ')
        }));

        if (debug) {
          await logger.logToStderr(`Admin consent granted for following resource ${permission.resourceId}, with delegated permissions: ${permission.scope.join(',')}`);
        }
      }

      permission.resourceAccess.filter(access => access.type === "Role").forEach(async (access: ResourceAccess) => {
        tasks.push(addRoleToServicePrincipal({
          objectId: sp.id,
          resourceId: permission.resourceId,
          appRoleId: access.id!
        }));

        if (debug) {
          await logger.logToStderr(`Admin consent granted for following resource ${permission.resourceId}, with application permission: ${access.id}`);
        }
      });
    });

    await Promise.all(tasks);
    return appInfo;
  },
  resolveApis: async ({ options, manifest, logger, verbose, debug }: {
    options: AppCreationOptions,
    manifest?: any,
    logger: Logger,
    verbose: boolean,
    debug: boolean
  }): Promise<RequiredResourceAccess[]> => {
    if (!options.apisDelegated && !options.apisApplication
      && (typeof manifest?.requiredResourceAccess === 'undefined' || manifest.requiredResourceAccess.length === 0)) {
      return [];
    }

    if (verbose) {
      await logger.logToStderr('Resolving requested APIs...');
    }

    const servicePrincipals = await odata.getAllItems<ServicePrincipalInfo>(`https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);

    let resolvedApis: RequiredResourceAccess[] = [];

    if (options.apisDelegated || options.apisApplication) {
      resolvedApis = await getRequiredResourceAccessForApis({
        servicePrincipals,
        apis: options.apisDelegated,
        scopeType: 'Scope',
        logger,
        debug
      });
      if (verbose) {
        await logger.logToStderr(`Resolved delegated permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
      }
      const resolvedApplicationApis = await getRequiredResourceAccessForApis({
        servicePrincipals,
        apis: options.apisApplication,
        scopeType: 'Role',
        logger,
        debug
      });
      if (verbose) {
        await logger.logToStderr(`Resolved application permissions: ${JSON.stringify(resolvedApplicationApis, null, 2)}`);
      }
      // merge resolved application APIs onto resolved delegated APIs
      resolvedApplicationApis.forEach(resolvedRequiredResource => {
        const requiredResource = resolvedApis.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
        if (requiredResource) {
          requiredResource.resourceAccess!.push(...resolvedRequiredResource.resourceAccess!);
        }
        else {
          resolvedApis.push(resolvedRequiredResource);
        }
      });
    }
    else {
      const manifestApis = (manifest.requiredResourceAccess as RequiredResourceAccess[]);

      manifestApis.forEach(manifestApi => {
        resolvedApis.push(manifestApi);

        const app = servicePrincipals.find(servicePrincipals => servicePrincipals.appId === manifestApi.resourceAppId);

        if (app) {
          manifestApi.resourceAccess!.forEach((res => {
            const resourceAccessPermission = {
              id: res.id,
              type: res.type
            };

            const oAuthValue = app.oauth2PermissionScopes.find(scp => scp.id === res.id)?.value;
            updateAppPermissions({
              spId: app.id,
              resourceAccessPermission,
              oAuth2PermissionValue: oAuthValue
            });
          }));
        }
      });
    }

    if (verbose) {
      await logger.logToStderr(`Merged delegated and application permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
      await logger.logToStderr(`App role assignments: ${JSON.stringify(entraApp.appPermissions.flatMap(permission => permission.resourceAccess.filter(access => access.type === "Role")), null, 2)}`);
      await logger.logToStderr(`OAuth2 permissions: ${JSON.stringify(entraApp.appPermissions.flatMap(permission => permission.scope), null, 2)}`);
    }

    return resolvedApis;
  },
  async getAppRegistrationByAppId(appId: string, properties?: string[]): Promise<Application> {
    let url = `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'`;

    if (properties) {
      url += `&$select=${properties.join(',')}`;
    }
    const apps = await odata.getAllItems<Application>(url);

    if (apps.length === 0) {
      throw `App with appId '${appId}' not found in Microsoft Entra ID`;
    }

    return apps[0];
  },
  async getAppRegistrationByAppName(appName: string, properties?: string[]): Promise<Application> {
    let url = `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'`;

    if (properties) {
      url += `&$select=${properties.join(',')}`;
    }

    const apps = await odata.getAllItems<Application>(url);

    if (apps.length === 0) {
      throw `App with name '${appName}' not found in Microsoft Entra ID`;
    }

    if (apps.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', apps);
      return await cli.handleMultipleResultsFound<Application>(`Multiple apps with name '${appName}' found in Microsoft Entra ID.`, resultAsKeyValuePair);
    }

    return apps[0];
  },
  async getAppRegistrationByObjectId(objectId: string, properties?: string[]): Promise<Application> {
    let url = `https://graph.microsoft.com/v1.0/applications/${objectId}`;

    if (properties) {
      url += `?$select=${properties.join(',')}`;
    }

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const app = await request.get<Application>(requestOptions);

    return app;
  }
};