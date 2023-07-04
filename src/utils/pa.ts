import { odata } from "./odata";

const paResource = 'https://api.powerapps.com';

export const pa = {
  /**
   * .
   * @param displayName the displayname of the app.
   */
  async getAppByDisplayName(displayName: string): Promise<any> {
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