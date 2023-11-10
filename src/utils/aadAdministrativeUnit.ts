import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import { odata } from "./odata.js";
import { formatting } from "./formatting.js";
import { Cli } from "../cli/Cli.js";

const graphResource = 'https://graph.microsoft.com';

export const aadAdministrativeUnit = {
  /**
   * Get an administrative unit by its display name.
   * @param displayName Administrative unit display name.
   * @returns The administrative unit.
   * @throws Error when administrative unit was not found.
   */
  async getAdministrativeUnitByDisplayName(displayName: string): Promise<AdministrativeUnit> {
    const administrativeUnits = await odata.getAllItems<AdministrativeUnit>(`${graphResource}/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (administrativeUnits.length === 0) {
      throw `The specified administrative unit '${displayName}' does not exist.`;
    }

    if (administrativeUnits.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', administrativeUnits);
      const selectedAdministrativeUnit = await Cli.handleMultipleResultsFound<AdministrativeUnit>(`Multiple administrative units with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedAdministrativeUnit;
    }

    return administrativeUnits[0];
  }
};