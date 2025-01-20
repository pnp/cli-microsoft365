import { cli } from "../cli/cli.js";
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
      return cli.handleMultipleResultsFound(`Multiple Power Page websites with name '${websiteName}' found`, resultAsKeyValuePair);
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
  }
};