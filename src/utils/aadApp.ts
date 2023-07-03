import { Application } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";

const graphResource = 'https://graph.microsoft.com';

export const aadApp = {
  /**
   * Retrieve a single app.
   * @param id App ID.
   */
  async getAppById(id: string): Promise<Application> {
    const requestOptionsForObjectId: CliRequestOptions = {
      url: `${graphResource}/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(id)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: Application[] }>(requestOptionsForObjectId);

    if (res.value.length === 0) {
      throw `No Azure AD application registration with ${id} found`;
    }

    return res.value[0];
  }
};