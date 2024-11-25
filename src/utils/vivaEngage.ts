import { cli } from '../cli/cli.js';
import { Community } from '../m365/viva/commands/engage/Community.js';
import request, { CliRequestOptions } from '../request.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export const vivaEngage = {
  /**
   * Get Viva Engage group ID by community ID.
   * @param communityId The ID of the Viva Engage community.
   * @returns The ID of the Viva Engage group.
   * @returns The Viva Engage community.
   */
  async getCommunityById(communityId: string, selectProperties: string[]): Promise<Community> {
    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}?$select=${selectProperties.join(',')}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const community = await request.get<Community>(requestOptions);

    if (!community) {
      throw `The specified Viva Engage community with ID '${communityId}' does not exist.`;
    }

    return community;
  },

  /**
   * Get Viva Engage community by display name.
   * @param displayName Community display name.
   * @param selectProperties Properties to select.
   * @returns The Viva Engage community.
   */
  async getCommunityByDisplayName(displayName: string, selectProperties: string[]): Promise<Community> {
    const communities = await odata.getAllItems<Community>(`https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=${selectProperties.join(',')}`);

    if (communities.length === 0) {
      throw `The specified Viva Engage community '${displayName}' does not exist.`;
    }

    if (communities.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', communities);
      const selectedCommunity = await cli.handleMultipleResultsFound<Community>(`Multiple Viva Engage communities with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedCommunity;
    }

    return communities[0];
  },

  /**
   * Get Viva Engage community by Microsoft Entra group ID.
   * Note: The Graph API doesn't support filtering by groupId, so we need to retrieve all communities and filter them in memory.
   * @param entraGroupId The ID of the Microsoft Entra group.
   * @param selectProperties Properties to select.
   * @returns The Viva Engage community.
   */
  async getCommunityByEntraGroupId(entraGroupId: string, selectProperties: string[]): Promise<Community> {
    const properties = selectProperties.includes('groupId') ? selectProperties : [...selectProperties, 'groupId'];
    const communities = await odata.getAllItems<Community>(`https://graph.microsoft.com/v1.0/employeeExperience/communities?$select=${properties.join(',')}`);

    const filteredCommunity = communities.find(c => c.groupId === entraGroupId);

    if (!filteredCommunity) {
      throw `The Microsoft Entra group with id '${entraGroupId}' is not associated with any Viva Engage community.`;
    }

    return filteredCommunity;
  }
};
