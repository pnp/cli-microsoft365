import { ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';

export const entraServicePrincipal = {
  async getServicePrincipalIdFromFromAppId(appId: string): Promise<string> {
    const apps = await odata.getAllItems<ServicePrincipal>(`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`);

    if (apps.length === 0) {
      throw `Service principal with appId '${appId}' not found in Microsoft Entra ID`;
    }

    return apps[0].id!;
  },
  async getServicePrincipalIdFromAppName(appName: string): Promise<string> {
    const apps = await odata.getAllItems<ServicePrincipal>(`https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`);

    if (apps.length === 0) {
      throw `Service principal with name '${appName}' not found in Microsoft Entra ID`;
    }

    if (apps.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', apps);
      return (await cli.handleMultipleResultsFound<ServicePrincipal>(`Multiple service principals with name '${appName}' found in Microsoft Entra ID.`, resultAsKeyValuePair)).id!;
    }

    return apps[0].id!;
  }
};