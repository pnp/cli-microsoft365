import { AdministrativeUnit } from '@microsoft/microsoft-graph-types';
import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';

export const entraAdministrativeUnit = {
  /**
   * Get an administrative unit by its display name.
   * @param displayName Administrative unit display name.
   * @param properties Properties to include in the response.
   * @returns The administrative unit.
   * @throws Error when administrative unit was not found.
   */
  async getAdministrativeUnitByDisplayName(displayName: string, properties?: string): Promise<AdministrativeUnit> {
    const queryParameters: string[] = [];

    if (properties) {
      const allProperties = properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    const graphResource = 'https://graph.microsoft.com';
    const administrativeUnits = await odata.getAllItems<AdministrativeUnit>(`${graphResource}/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'${queryString}`);

    if (administrativeUnits.length === 0) {
      throw `The specified administrative unit '${displayName}' does not exist.`;
    }

    if (administrativeUnits.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', administrativeUnits);
      const selectedAdministrativeUnit = await cli.handleMultipleResultsFound<AdministrativeUnit>(`Multiple administrative units with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedAdministrativeUnit;
    }

    return administrativeUnits[0];
  }
};