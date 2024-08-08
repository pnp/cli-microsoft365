import { Logger } from "../cli/Logger.js";
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from "../request.js";

export const drive = {
  /**
   * Retrieves the Drive associated with the specified site and folder URL.
   * @param siteId Site ID
   * @param folderUrl Folder URL
   * @param logger The logger object.
   * @param verbose Set for verbose logging
   * @returns The Drive associated with the folder URL.
   */
  async getDrive(siteId: string, folderUrl: URL, logger?: Logger, verbose?: boolean): Promise<Drive> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the drive associated with the folder URL: ${folderUrl}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const drives = await request.get<{ value: Drive[] }>(requestOptions);
    const lowerCaseFolderUrl: string = folderUrl.href.toLowerCase();

    const drive: Drive | undefined = drives.value
      .sort((a, b) => (b.webUrl as string).localeCompare(a.webUrl as string))
      .find((d: Drive) => {
        const driveUrl: string = (d.webUrl as string).toLowerCase();

        return lowerCaseFolderUrl.startsWith(driveUrl) &&
          (driveUrl.length === lowerCaseFolderUrl.length ||
            lowerCaseFolderUrl[driveUrl.length] === '/');
      });

    if (!drive) {
      throw `Drive '${folderUrl.href}' not found`;
    }

    return drive;
  },

  /**
   * Retrieves the ID of a drive item (file, folder, etc.) associated with the given drive and item URL.
   * @param drive The Drive object containing the item
   * @param itemUrl Item URL
   * @param logger The logger object.
   * @param verbose Set for verbose logging
   * @returns Drive item ID
   */
  async getDriveItemId(drive: Drive, itemUrl: URL, logger?: Logger, verbose?: boolean): Promise<string> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving ID of a drive item associated with item URL: ${itemUrl}...`);
    }

    const relativeItemUrl: string = itemUrl.href.replace(new RegExp(`${drive.webUrl}`, 'i'), '').replace(/\/+$/, '');

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/drives/${drive.id}/root${relativeItemUrl ? `:${relativeItemUrl}` : ''}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const driveItem = await request.get<DriveItem>(requestOptions);
    return driveItem?.id as string;
  }
};