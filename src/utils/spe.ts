import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import { odata } from './odata.js';

export interface SpeContainerType {
  id: string;
  name: string;
  owningAppId: string;
  billingClassification: string;
  createdDateTime: string;
  expirationDateTime: string;
  etag: string;
  settings: {
    urlTemplate: string;
    isDiscoverabilityEnabled: boolean;
    isSearchEnabled: boolean;
    isItemVersioningEnabled: boolean;
    itemMajorVersionLimit: number;
    maxStoragePerContainerInBytes: number;
    isSharingRestricted: boolean;
    consumingTenantOverridables: string;
  }
}

export interface SpeContainer {
  id: string;
  displayName: string;
  containerTypeId: string;
  createdDateTime: string;
}

const graphResource = 'https://graph.microsoft.com';

export const spe = {
  /**
   * Get the ID of a container type by its name.
   * @param name Name of the container type to search for
   * @returns ID of the container type
   */
  async getContainerTypeIdByName(name: string): Promise<string> {
    const containerTypes = await odata.getAllItems<SpeContainerType>(`${graphResource}/beta/storage/fileStorage/containerTypes?$select=id,name&$filter=name eq '${formatting.encodeQueryParameter(name)}'`);

    if (containerTypes.length === 0) {
      throw new Error(`The specified container type '${name}' does not exist.`);
    }

    if (containerTypes.length > 1) {
      const containerTypeKeyValuePair = formatting.convertArrayToHashTable('id', containerTypes);
      const containerType = await cli.handleMultipleResultsFound<SpeContainerType>(`Multiple container types with name '${name}' found.`, containerTypeKeyValuePair);
      return containerType.id;
    }

    return containerTypes[0].id;
  },

  /**
   * Get the ID of a container by its name.
   * @param containerTypeId ID of the container type.
   * @param name Name of the container to search for.
   * @returns ID of the container.
   */
  async getContainerIdByNameAndContainerTypeId(containerTypeId: string, name: string): Promise<string> {
    const containers = await odata.getAllItems<SpeContainer>(`${graphResource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`);
    const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === name.toLowerCase());

    if (matchingContainers.length === 0) {
      throw new Error(`The specified container '${name}' does not exist.`);
    }

    if (matchingContainers.length > 1) {
      const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
      const container = await cli.handleMultipleResultsFound<SpeContainer>(`Multiple containers with name '${name}' found.`, containerKeyValuePair);
      return container.id;
    }

    return matchingContainers[0].id;
  },

  /**
   * Get the ID of a container by its display name.
   * @param name Name of the container to search for.
   * @returns ID of the container.
   */
  async getContainerIdByName(name: string): Promise<string> {
    const containers = await odata.getAllItems<SpeContainer>(`${graphResource}/v1.0/storage/fileStorage/containers?$select=id,displayName`);
    const matchingContainers = containers.filter(c => c.displayName === name);

    if (matchingContainers.length === 0) {
      throw new Error(`The specified container '${name}' does not exist.`);
    }

    if (matchingContainers.length > 1) {
      const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
      const container = await cli.handleMultipleResultsFound<SpeContainer>(`Multiple containers with name '${name}' found.`, containerKeyValuePair);
      return container.id;
    }

    return matchingContainers[0].id;
  }
};
