import { Channel, Team } from '@microsoft/microsoft-graph-types';
import { CliRequestOptions } from '../request.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';
import { cli } from '../cli/cli.js';

const graphResource = 'https://graph.microsoft.com';

export const teams = {

  /**
   * Retrieve the team by its name.  
   * @param displayName Name of the team to retrieve.
   * @throws Error if the team cannot be found.
   * @throws Error when multiple teams with the same name and prompting is disabled.
   * @returns The team.
   */
  async getTeamByDisplayName(displayName: string): Promise<Team> {
    const teams = await odata.getAllItems<Team>(`${graphResource}/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (!teams.length) {
      throw Error(`The specified team '${displayName}' does not exist.`);
    }

    if (teams.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', teams);
      const result = await cli.handleMultipleResultsFound<Team>(`Multiple teams with name '${displayName}' found.`, resultAsKeyValuePair);
      return result;
    }

    return teams[0];
  },

  /**
   * Retrieve the id of a team by its name.  
   * @param displayName Name of the team to retrieve.
   * @throws Error if the team cannot be found.
   * @throws Error when multiple teams with the same name and prompting is disabled.
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
     * Retrieves the channel by its name in a Microsoft Teams team.
     * @param teamId The ID of the team.
     * @param name The name of the channel.
     * @returns The channel.
     * @throws Throws an error if the specified channel does not exist in the team.
     */
  async getChannelByDisplayName(teamId: string, name: string): Promise<Channel> {
    const channelRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/teams/${teamId}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await odata.getAllItems<Channel>(channelRequestOptions);
    // Only one channel can have the same name in a team
    const channelItem = response[0];

    if (!channelItem) {
      throw Error(`The channel '${name}' does not exist in the Microsoft Teams team with ID '${teamId}'.`);
    }

    return channelItem;
  },

  /**
   * Retrieves the channel ID by its name in a Microsoft Teams team.
   * @param teamId The ID of the team.
   * @param name The name of the channel.
   * @returns The ID of the channel.
   * @throws Throws an error if the specified channel does not exist in the team.
   */
  async getChannelIdByDisplayName(teamId: string, name: string): Promise<string> {
    const channelRequestOptions: CliRequestOptions = {
      url: `${graphResource}/v1.0/teams/${teamId}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await odata.getAllItems<Channel>(channelRequestOptions);
    // Only one channel can have the same name in a team
    const channelItem = response[0];

    if (!channelItem) {
      throw Error(`The channel '${name}' does not exist in the Microsoft Teams team with ID '${teamId}'.`);
    }

    return channelItem.id!;
  }
};