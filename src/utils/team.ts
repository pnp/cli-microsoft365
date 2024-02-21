import { Channel, Team } from "@microsoft/microsoft-graph-types";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";
import { cli } from "../cli/cli.js";

const graphResource = 'https://graph.microsoft.com';

export const team = {

  /**
   * Retrieves the team ID based on the provided ID or display name.
   * If the ID is provided, it checks if the team exists.
   * If the display name is provided, it retrieves the team ID by the display name.
   * @param id The ID of the team to retrieve. Optional.
   * @param name The display name of the team to retrieve. Optional.
   * @returns The ID of the team.
   */
  async getTeamId(id?: string, name?: string): Promise<string> {
    if (id) {
      return this.verifyTeamExistsById(id);
    }

    return this.getTeamIdByDisplayName(name!);
  },

  /**
   * Verifies if a team exists based on its ID.
   * @param id The ID of the team to verify.
   * @returns The ID of the team if it exists.
   */
  async verifyTeamExistsById(id: string): Promise<string> {
    const teamRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/teams/${id}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    const response = await request.get<{ id: string }>(teamRequestOptions);
    return response.id;
  },

  /**
   * Retrieve the id of a team by the name.
   * @param displayName Name of the team to retrieve.
   * @throws Error if the team cannot be found.
   * @throws Error when multiple teams with the same name were found.
   * @returns The ID of the team.
   */
  async getTeamIdByDisplayName(displayName: string): Promise<string> {
    const teams = await odata.getAllItems<Team>(`${graphResource}/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`);

    if (!teams.length) {
      throw Error(`The specified team '${displayName}' does not exist.`);
    }

    if (teams.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', teams);
      const result = await cli.handleMultipleResultsFound<Team>(`Multiple teams with name '${displayName}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return teams[0].id!;
  },

  /**
   * Retrieves the channel ID based on the provided team ID, channel ID, or channel name.
   * If the channel ID is provided, it verifies that the channel exists in the team.
   * If the channel name is provided, it retrieves the channel ID by name.
   * @param teamId The ID of the team.
   * @param channelId The ID of the channel (optional).
   * @param channelName The name of the channel (optional).
   * @returns The ID of the channel.
   */
  async getChannelId(teamId: string, channelId?: string, channelName?: string): Promise<string> {
    if (channelId) {
      return this.verifyChannelExistsById(teamId, channelId);
    }

    return this.getChannelIdByName(teamId, channelName!);
  },

  /**
   * Verifies if a channel exists in a Microsoft Teams team.
   * @param teamId The ID of the team.
   * @param channelId The ID of the channel.
   * @returns The ID of the channel if it exists.
   * @throws Throws an error if the specified channel does not exist in the team.
   */
  async verifyChannelExistsById(teamId: string, channelId: string): Promise<string> {
    const channelRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/teams/${teamId}/channels/${channelId}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      await request.get<{ id: string }>(channelRequestOptions);
    }
    catch (err: any) {
      throw Error('The specified channel does not exist in the Microsoft Teams team.');
    }

    return channelId;
  },


  /**
   * Retrieves the channel ID by its name in a Microsoft Teams team.
   * @param teamId The ID of the team.
   * @param name The name of the channel.
   * @returns The ID of the channel.
   * @throws Throws an error if the specified channel does not exist in the team.
   */
  async getChannelIdByName(teamId: string, name: string): Promise<string> {
    const channelRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/teams/${teamId}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    // Only one channel can have the same name in a team
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw Error('The specified channel does not exist in the Microsoft Teams team');
    }

    return channelItem.id!;
  }
};