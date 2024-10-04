import { cli } from '../cli/cli.js';
import { Community } from '../m365/viva/commands/engage/Community.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export const vivaEngage = {
  /**
   * Get Viva Engage community ID by display name.
   * @param displayName Community display name.
   * @returns The ID of the Viva Engage community.
   * @throws Error when the community was not found.
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
   * Get Viva Engage community ID by Entra group ID.
   * @param entraGroupId The Microsoft Entra group ID.
   * @returns The ID of the Viva Engage community.
   * @throws Error when the community was not found.
   */
  async getCommunityIdByEntraGroupId(entraGroupId: string): Promise<string> {
    // The Graph API doesn't support filtering by groupId
    const communities = await odata.getAllItems<Community>('https://graph.microsoft.com/v1.0/employeeExperience/communities?$select=id,groupId');

    const filtereCommunities = communities.filter(c => c.groupId === entraGroupId);

    if (filtereCommunities.length === 0) {
      throw `The Microsoft Entra group with id '${entraGroupId}' is not associated with any Viva Engage community.`;
    }

    return filtereCommunities[0].id;
  }
};