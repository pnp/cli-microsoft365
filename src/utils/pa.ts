import { Logger } from "../cli/Logger";
import { odata } from "./odata";

const paResource = 'https://api.powerapps.com';

export const pa = {
  /**
   * Get a Power App by dusplay name
   * Returns the Power App
   * @param displayName The displayname of the app.
   * @param logger The logger object
   * @param verbose Set for verbose logging
   */
  async getAppByDisplayName(displayName: string, logger?: Logger, verbose?: boolean): Promise<any> {
    if (verbose && logger) {
      logger.logToStderr(`Retrieving the Power App with displayname ${displayName}`);
    }
    const url = `${paResource}/providers/Microsoft.PowerApps/apps?api-version=2017-08-01`;

    const apps: any = await odata.getAllItems<{ name: string; displayName: string; properties: { displayName: string } }>(url);

    if (apps.length === 0) {
      throw 'No apps found';
    }

    const app = apps.find((a: any) => {
      return a.properties.displayName.toLowerCase() === `${displayName}`.toLowerCase();
    });

    if (!app) {
      throw `No app found with displayName '${displayName}'`;
    }

    return app;
  }
};