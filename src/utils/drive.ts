
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from "../request.js";
import { Logger } from "../cli/Logger.js";

export const drive = {
  /**
   * Retrieves the Drive associated with the specified site and URL.
   * @param siteId Site ID
   * @param url Drive URL
   * @param logger The logger object
   * @param verbose Set for verbose logging
   * @returns The Drive associated with the drive URL.
   */
  async getDriveByUrl(siteId: string, url: URL, logger?: Logger, verbose?: boolean): Promise<Drive> {
    if (verbose && logger) {
      await logger.logToStderr(`Retrieving drive information for URL: ${url.href}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const drives = await request.get<{ value: Drive[] }>(requestOptions);

    const lowerCaseFolderUrl: string = url.href.toLowerCase();

    const drive: Drive | undefined = drives.value
      .sort((a, b) => (b.webUrl as string).localeCompare(a.webUrl as string))
      .find((d: Drive) => {
        const driveUrl: string = (d.webUrl as string).toLowerCase();

        return lowerCaseFolderUrl.startsWith(driveUrl) &&
          (driveUrl.length === lowerCaseFolderUrl.length ||
            lowerCaseFolderUrl[driveUrl.length] === '/');
      });

    if (!drive) {
      throw new Error(`Drive '${url.href}' not found`);
    }

    return drive;
  },

  /**
   * Retrieves the ID of a drive item (file, folder, etc.) associated with the given drive and item URL.
   * @param drive The Drive object containing the item
   * @param itemUrl Item URL
   * @param logger The logger object
   * @param verbose Set for verbose logging
   * @returns Drive item ID
   */
  async getDriveItemId(drive: Drive, itemUrl: URL, logger?: Logger, verbose?: boolean): Promise<string> {
    const relativeItemUrl: string = itemUrl.href.replace(new RegExp(`${drive.webUrl}`, 'i'), '').replace(/\/+$/, '');

    if (verbose && logger) {
      await logger.logToStderr(`Retrieving drive item ID for URL: ${relativeItemUrl}`);
    }

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