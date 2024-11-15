import { UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';
import { cli } from '../cli/cli.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';
import request, { CliRequestOptions } from '../request.js';

export const roleDefinition = {
  /**
   * Get an Entra ID (directory) role by its name
   * @param displayName Role definition display name.
   * @param properties Comma-separated list of properties to include in the response.
   * @returns The role definition.
   * @throws Error when role definition was not found.
   */
  async getRoleDefinitionByDisplayName(displayName: string, properties?: string): Promise<UnifiedRoleDefinition> {
    let url = `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
    if (properties) {
      url += `&$select=${properties}`;
    }
    const roleDefinitions = await odata.getAllItems<UnifiedRoleDefinition>(url);

    if (roleDefinitions.length === 0) {
      throw `The specified role definition '${displayName}' does not exist.`;
    }

    if (roleDefinitions.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', roleDefinitions);
      const selectedRoleDefinition = await cli.handleMultipleResultsFound<UnifiedRoleDefinition>(`Multiple role definitions with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedRoleDefinition;
    }

    return roleDefinitions[0];
  },

  /**
   * Get an Entra ID (directory) role by its id
   * @param id Role definition id.
   * @param properties Comma-separated list of properties to include in the response.
   * @returns The role definition.
   * @throws Error when role definition was not found.
   */
  async getRoleDefinitionById(id: string, properties?: string): Promise<UnifiedRoleDefinition> {
    let url = `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${id}`;
    if (properties) {
      url += `?$select=${properties}`;
    }

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<UnifiedRoleDefinition>(requestOptions);
  }
};