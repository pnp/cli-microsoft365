import { RoleDefinition } from "@microsoft/microsoft-graph-types";
import { odata } from "./odata.js";
import { formatting } from "./formatting.js";
import { Cli } from "../cli/Cli.js";

const graphResource = 'https://graph.microsoft.com';

export const roleDefinition = {
  /**
   * 
   * @param displayName Role definition display name.
   * @returns The role definition.
   * @throws Error when role definition was not found.
   */
  async getRoleDefinitionByDisplayName(displayName: string): Promise<RoleDefinition> {
    const roleDefinitions = await odata.getAllItems<RoleDefinition>(`${graphResource}/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (roleDefinitions.length === 0) {
      throw `The specified role definition '${displayName}' does not exist.`;
    }

    if (roleDefinitions.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', roleDefinitions);
      const selectedRoleDefinition = await Cli.handleMultipleResultsFound<RoleDefinition>(`Multiple role definitions with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedRoleDefinition;
    }

    return roleDefinitions[0];
  }
};