import { cli } from '../cli/cli.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export const vivaEngageRole = {
  async getRoleIdByName(roleName: string): Promise<string> {
    // the endpoint doesn't support filtering by displayName
    const response = await odata.getAllItems<{ id: string; displayName: string }>(`https://graph.microsoft.com/beta/employeeExperience/roles`);
    const roles = response.filter(role => role.displayName.toLowerCase() === roleName.toLowerCase());

    if (roles.length === 0) {
      throw `The specified Viva Engage role '${roleName}' does not exist.`;
    }
    if (roles.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', roles);
      const selectedRole = await cli.handleMultipleResultsFound<{ id: string; displayName: string }>(`Multiple Viva Engage roles with name '${roleName}' found.`, resultAsKeyValuePair);
      return selectedRole.id;
    }
    return roles[0].id;
  }
};