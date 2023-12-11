import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import { odata } from "./odata.js";
import { formatting } from "./formatting.js";
import { cli } from "../cli/cli.js";
import { aadUser } from "./aadUser.js";
import { aadGroup } from "./aadGroup.js";
import { aadDevice } from "./aadDevice.js";

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
      const selectedAdministrativeUnit = await cli.handleMultipleResultsFound<AdministrativeUnit>(`Multiple administrative units with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedAdministrativeUnit;
    }

    return administrativeUnits[0];
  },

  async getMemberIdByName(name: string): Promise<string> {
    let userId;
    let groupId;
    let deviceId;
    const objectIds: any[] = [];
    try {
      userId = await aadUser.getUserIdByUpn(name);

      // is not possible that group name or device name is same as user UPN
      if (userId) {
        return userId;
      }
    }
    catch { }

    try {
      groupId = await aadGroup.getGroupIdByDisplayName(name);
      objectIds.push({ id: groupId });
    }
    catch { }

    try {
      deviceId = (await aadDevice.getDeviceByDisplayName(name)).id!;
      objectIds.push({ id: deviceId });
    }
    catch { }

    if (objectIds.length === 0) {
      throw Error(`The specified member '${name}' does not exist.`);
    }

    if (objectIds.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', objectIds);
      const selectedMember = await cli.handleMultipleResultsFound<any>(`Multiple members with name '${name}' found.`, resultAsKeyValuePair);
      return selectedMember.id;
    }

    return objectIds[0].id;
  }
};