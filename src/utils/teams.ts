import { Team } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request";
import { aadGroup } from "./aadGroup";
import { formatting } from "./formatting";
import { Logger } from "../cli/Logger";

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