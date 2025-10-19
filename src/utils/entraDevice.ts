import { Device } from "@microsoft/microsoft-graph-types";
import { odata } from "./odata.js";
import { formatting } from "./formatting.js";
import { cli } from "../cli/cli.js";

const graphResource = 'https://graph.microsoft.com';

export const entraDevice = {
  /**
   * Get a device by its display name.
   * @param displayName Device display name.
   * @returns The device.
   * @throws Error when device was not found.
   */
  async getDeviceByDisplayName(displayName: string): Promise<Device> {
    const devices = await odata.getAllItems<Device>(`${graphResource}/v1.0/devices?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`);

    if (devices.length === 0) {
      throw new Error(`The specified device '${displayName}' does not exist.`);
    }

    if (devices.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', devices);
      const selectedDevice = await cli.handleMultipleResultsFound<Device>(`Multiple devices with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedDevice;
    }

    return devices[0];
  }
};