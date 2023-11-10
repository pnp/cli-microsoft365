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
      return response.properties.linkedEnvironmentMetadata.instanceApiUrl;
    }
    catch (ex: any) {
      throw Error(`The environment '${environment}' could not be retrieved. See the inner exception for more details: ${ex.message}`);
    }
  }
};