import { Logger } from "../cli/Logger.js";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";

const powerPlatformResource = 'https://api.bap.microsoft.com';

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

      return Promise.resolve(response.properties.linkedEnvironmentMetadata.instanceApiUrl);
    }
    catch (ex: any) {
      throw Error(`The environment '${environment}' could not be retrieved. See the inner exception for more details: ${ex.message}`);
    }
  },

  /**
 * Get a solution publisher by name
 * Returns solution publisher object
 * @param dynamicsApiUrl The dynamics api url of the environment
 * @param name The name of the solution publisher.
 * @param logger The logger object
 * @param verbose Set for verbose logging
 */
  async getSolutionPublisherByName(dynamicsApiUrl: string, name: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the solution publisher by name ${name}`);
    }
    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.0/publishers?$filter=friendlyname eq \'${name}\'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw Error(`The specified solution publisher '${name}' does not exist.`);
    }

    if (result.value.length > 1) {
      throw Error(`Multiple solution publishers with name '${name}' found: ${result.value.map(x => x.publisherid).join(',')}`);
    }

    return result.value[0];
  }
};