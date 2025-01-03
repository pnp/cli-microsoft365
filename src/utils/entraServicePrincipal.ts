import { ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';

export const entraServicePrincipal = {
  /**
   * Get service principal by its appId
   * @param appId App id.
   * @param properties Comma-separated list of properties to include in the response.
   * @returns The service principal.
   * @throws Error when service principal was not found.
   */
  async getServicePrincipalFromAppId(appId: string, properties?: string): Promise<ServicePrincipal> {
    let url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'`;

    if (properties) {
      url += `&$select=${properties}`;
    }

    const apps = await odata.getAllItems<ServicePrincipal>(url);

    if (apps.length === 0) {
      throw `Service principal with appId '${appId}' not found in Microsoft Entra ID`;
    }

    return apps[0];
  },
  /**
   * Get service principal by its name
   * @param appName Service principal name.
   * @param properties Comma-separated list of properties to include in the response.
   * @returns The service principal.
   * @throws Error when service principal was not found.
   */
  async getServicePrincipalFromAppName(appName: string, properties?: string): Promise<ServicePrincipal> {
    let url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'`;

    if (properties) {
      url += `&$select=${properties}`;
    }

    const apps = await odata.getAllItems<ServicePrincipal>(url);

    if (apps.length === 0) {
      throw `Service principal with name '${appName}' not found in Microsoft Entra ID`;
    }

    if (apps.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', apps);
      return await cli.handleMultipleResultsFound<ServicePrincipal>(`Multiple service principals with name '${appName}' found in Microsoft Entra ID.`, resultAsKeyValuePair);
    }

    return apps[0];
  }
};