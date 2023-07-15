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
 * Get a solution by name
 * Returns the solution object
 * @param dynamicsApiUrl The dynamics api url of the environment
 * @param name The name of the solution
 * @param logger The logger object
 * @param verbose Set for verbose logging
 */
  async getSolutionByName(dynamicsApiUrl: string, name: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the solution by name ${name}`);
    }
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
      throw Error(`Multiple solutions with name '${name}' found: ${result.value.map(x => x.solutionid).join(',')}`);
    }

    return result.value[0];
  }
};