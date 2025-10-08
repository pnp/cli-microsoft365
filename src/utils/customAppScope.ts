import { Entity, NullableOption, UnifiedRoleDefinition } from '@microsoft/microsoft-graph-types';
import { cli } from '../cli/cli.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export interface CustomAppScope extends Entity {
  // The display name of the app-specific resource represented by the app scope
  displayName?: NullableOption<string>;
  // The type of app-specific resource represented by the app scope
  type?: NullableOption<string>;
  // An open dictionary type that holds workload-specific properties for the scope object
  customAttributes?: NullableOption<customAppScopeAttributesDictionary>;
}

export interface customAppScopeAttributesDictionary {
  // Indicates whether the object is an exclusive scope
  exclusive?: NullableOption<boolean>;
  // A filter query that defines how you segment your recipients that admins can manage
  recipientFilter?: NullableOption<string>;
}

export const customAppScope = {
  /**
   * Get a custom application scope by its name
   * @param displayName Custom application scope display name.
   * @param properties Comma-separated list of properties to include in the response.
   * @returns The custom application scope.
   * @throws Error when role definition was not found.
   */
  async getCustomAppScopeByDisplayName(displayName: string, properties?: string): Promise<UnifiedRoleDefinition> {
    let url = `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`;

    if (properties) {
      url += `&$select=${properties}`;
    }

    const customAppScopes = await odata.getAllItems<CustomAppScope>(url);

    if (customAppScopes.length === 0) {
      throw new Error(`The specified custom application scope '${displayName}' does not exist.`);
    }

    if (customAppScopes.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', customAppScopes);
      const selectedCustomAppScope = await cli.handleMultipleResultsFound<CustomAppScope>(`Multiple custom application scopes with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedCustomAppScope;
    }

    return customAppScopes[0];
  }
};