import { ExtensionProperty } from '@microsoft/microsoft-graph-types';
import { formatting } from './formatting.js';
import { odata } from './odata.js';

export const directoryExtension = {
  /**
   * Get a directory extension by its name registered for an application.
   * @param name Role definition display name.
   * @param appObjectId Application object id.
   * @param properties List of properties to include in the response.
   * @returns The directory extensions.
   * @throws Error when directory extension was not found.
   */
  async getDirectoryExtensionByName(name: string, appObjectId: string, properties?: string[]): Promise<ExtensionProperty> {
    let url = `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties?$filter=name eq '${formatting.encodeQueryParameter(name)}'`;

    if (properties) {
      url += `&$select=${properties.join(',')}`;
    }

    const extensionProperties = await odata.getAllItems<ExtensionProperty>(url);

    if (extensionProperties.length === 0) {
      throw `The specified directory extension '${name}' does not exist.`;
    }

    // there can be only one directory extension with a given name
    return extensionProperties[0];
  }
};
