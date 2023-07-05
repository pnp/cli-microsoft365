import { cli } from "../cli/cli.js";
import { Logger } from "../cli/Logger.js";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";

const powerPlatformResource = 'https://api.bap.microsoft.com';

export interface PowerPageWebsite {
  id: string,
  name: string,
  createdOn: string,
  templateName: string,
  websiteUrl: string,
  tenantId: string,
  dataverseInstanceUrl: string,
  environmentName: string,
  environmentId: string,
  dataverseOrganizationId: string,
  selectedBaseLanguage: number,
  customHostNames: Array<string>,
  websiteRecordId: string,
  subdomain: string,
  packageInstallStatus: string,
  type: string,
  trialExpiringInDays: number,
  suspendedWebsiteDeletingInDays: number,
  packageVersion: string,
  isEarlyUpgradeEnabled: boolean,
  isCustomErrorEnabled: boolean,
  applicationUserAadAppId: string,
  ownerId: string,
  status: string,
  siteVisibility: string,
  dataModel: string
}

export const powerPlatform = {
  async getDynamicsInstanceApiUrl(environment: string, asAdmin?: boolean): Promise<string> {
    let url: string = '';
    if (asAdmin) {
      url = `${powerPlatformResource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${formatting.encodeQueryParameter(environment)}`;
    }
    else {
      url = `${powerPlatformResource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(environment)}`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${url}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<any>(requestOptions);
      return response.properties.linkedEnvironmentMetadata.instanceApiUrl;
    }
    catch (ex: any) {
      throw Error(`The environment '${environment}' could not be retrieved. See the inner exception for more details: ${ex.message}`);
    }
  },

  async getWebsiteById(environment: string, id: string): Promise<PowerPageWebsite> {
    const requestOptions: CliRequestOptions = {
      url: `https://api.powerplatform.com/powerpages/environments/${environment}/websites/${id}?api-version=2022-03-01-preview`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<PowerPageWebsite>(requestOptions);
      return response;
    }
    catch (ex: any) {
      throw Error(`The specified Power Page website with id '${id}' does not exist.`);
    }
  },

  async getWebsiteByName(environment: string, websiteName: string): Promise<PowerPageWebsite> {
    const response = await odata.getAllItems<PowerPageWebsite>(`https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`);

    const items = response.filter(response => response.name === websiteName);

    if (items.length === 0) {
      throw Error(`The specified Power Page website '${websiteName}' does not exist.`);
    }

    if (items.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('websiteUrl', items);
      return cli.handleMultipleResultsFound(`Multiple Power Page websites with name '${websiteName}' found.`, resultAsKeyValuePair);
    }

    return items[0];
  },

  async getWebsiteByUrl(environment: string, url: string): Promise<PowerPageWebsite> {
    const response = await odata.getAllItems<PowerPageWebsite>(`https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`);

    const items = response.filter(response => response.websiteUrl === url);

    if (items.length === 0) {
      throw Error(`The specified Power Page website with url '${url}' does not exist.`);
    }

    return items[0];
  },

  /**
   * Get a card by name
   * Returns a card object
   * @param dynamicsApiUrl The dynamics api url of the environment
   * @param name The name of the card
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async getCardByName(dynamicsApiUrl: string, name: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving the card with name ${name}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.1/cards?$filter=name eq '${name}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw Error(`The specified card '${name}' does not exist.`);
    }

    if (result.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('cardid', result.value);
      return cli.handleMultipleResultsFound(`Multiple cards with name '${name}' found.`, resultAsKeyValuePair);
    }

    return result.value[0];
  },

  /**
 * Get a solution by name
 * Returns the solution object
 * @param dynamicsApiUrl The dynamics api url of the environment
 * @param name The name of the solution
 */
  async getSolutionByName(dynamicsApiUrl: string, name: string): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq \'${name}\'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw Error(`The specified solution '${name}' does not exist.`);
    }

    if (result.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('solutionid', result.value);
      return cli.handleMultipleResultsFound(`Multiple solutions with name '${name}' found.`, resultAsKeyValuePair);
    }

    return result.value[0];
  }
};