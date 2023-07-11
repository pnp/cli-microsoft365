import { Team } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request";
import { aadGroup } from "./aadGroup";
import { formatting } from "./formatting";

const teamsResource = 'https://graph.microsoft.com';

/**
* Retrieves the team
* @param name The name of the team 
*/
export const teams = {
  async getTeamByName(name: string): Promise<Team> {
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