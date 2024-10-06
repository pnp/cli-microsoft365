import { cli } from '../cli/cli.js';
import { Community } from '../m365/viva/commands/engage/Community.js';
import request, { CliRequestOptions } from '../request.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export const vivaEngage = {
  /**
   * Get Viva Engage community ID by display name.
   * @param displayName Community display name.
   * @returns The ID of the Viva Engage community.
   */
  async getCommunityIdByDisplayName(displayName: string): Promise<string> {
    const communities = await odata.getAllItems<Community>(`https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (communities.length === 0) {
      throw `The specified Viva Engage community '${displayName}' does not exist.`;
    }

    if (communities.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', communities);
      const selectedCommunity = await cli.handleMultipleResultsFound<Community>(`Multiple Viva Engage communities with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedCommunity.id;
    }

    return communities[0].id;
  },

  /**
   * Get Viva Engage community ID by Microsoft Entra group ID.
   * Note: The Graph API doesn't support filtering by groupId, so we need to retrieve all communities and filter them in memory.
   * @param entraGroupId The ID of the Microsoft Entra group.
   * @returns The ID of the Viva Engage community.
   */
  async getCommunityIdByEntraGroupId(entraGroupId: string): Promise<string> {
    const communities = await odata.getAllItems<Community>('https://graph.microsoft.com/v1.0/employeeExperience/communities?$select=id,groupId');

    const filtereCommunities = communities.filter(c => c.groupId === entraGroupId);

    if (filtereCommunities.length === 0) {
      throw `The Microsoft Entra group with id '${entraGroupId}' is not associated with any Viva Engage community.`;
    }

    return filtereCommunities[0].id;
  },

  /**
   * Get Viva Engage group ID by community ID.
   * @param communityId The ID of the Viva Engage community.
   * @returns The ID of the Viva Engage group.
   */
  async getEntraGroupIdByCommunityId(communityId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}?$select=groupId`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const community = await request.get<Community>(requestOptions);

    if (!community) {
      throw `The specified Viva Engage community with ID '${communityId}' does not exist.`;
    }

    return community.groupId;
  },

  /**
   * Get Viva Engage group ID by community display name.
   * @param displayName Community display name.
   * @returns The ID of the Viva Engage group.
   */
  async getEntraGroupIdByCommunityDisplayName(displayName: string): Promise<string> {
    const communityId = await this.getCommunityIdByDisplayName(displayName);
    return await this.getEntraGroupIdByCommunityId(communityId);
  }
};