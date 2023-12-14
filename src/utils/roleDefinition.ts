import { RoleDefinition } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';

export const roleDefinition = {
  /**
   * Get a directory (Microsoft Entra) role
   * @param displayName Role definition display name.
   * @returns The role definition.
   * @throws Error when role definition was not found.
   */
  async getRoleDefinitionByDisplayName(displayName: string): Promise<RoleDefinition> {
    const graphResource = 'https://graph.microsoft.com';

    const roleDefinitions = await odata.getAllItems<RoleDefinition>(`${graphResource}/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (roleDefinitions.length === 0) {
      throw `The specified role definition '${displayName}' does not exist.`;
    }

    if (roleDefinitions.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', roleDefinitions);
      const selectedRoleDefinition = await cli.handleMultipleResultsFound<RoleDefinition>(`Multiple role definitions with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedRoleDefinition;
    }

    return roleDefinitions[0];
  }
};