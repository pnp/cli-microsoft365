import { Team } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { aadGroup } from "./aadGroup.js";
import { formatting } from "./formatting.js";
import { Logger } from "../cli/Logger.js";

const teamsResource = 'https://graph.microsoft.com';

/**
* Retrieves the team
* Returns a team object
* @param name The name of the team 
* @param logger The logger object
* @param verbose Set for verbose logging
*/
export const teams = {
  async getTeamByName(name: string, logger?: Logger, verbose?: boolean): Promise<Team> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the team by name ${name}`);
    }

    const groupId: string = await aadGroup.getGroupIdByDisplayName(name);

    const requestOptions: CliRequestOptions = {
      url: `${teamsResource}/v1.0/teams/${formatting.encodeQueryParameter(groupId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<Team>(requestOptions);
  }
};