import { Application } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";
import { Logger } from "../cli/Logger.js";

const graphResource = 'https://graph.microsoft.com';

export const aadApp = {
  /**
   * Retrieve a single app.
   * Returns an Application object
   * @param id App ID.
   */
  async getAppById(id: string, logger?: Logger, verbose?: boolean): Promise<Application> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the app with id ${id}`);
    }
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