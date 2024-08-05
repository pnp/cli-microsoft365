import auth from "../Auth.js";
import { CommandError } from "../Command.js";

export const accessToken = {
  isAppOnlyAccessToken(accessToken: string): boolean | undefined {
    let isAppOnlyAccessToken: boolean | undefined;

    if (!accessToken || accessToken.length === 0) {
      return isAppOnlyAccessToken;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return isAppOnlyAccessToken;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      isAppOnlyAccessToken = token.idtyp === 'app';
    }
    catch {
    }

    return isAppOnlyAccessToken;
  },

  getTenantIdFromAccessToken(accessToken: string): string {
    let tenantId: string = '';

    if (!accessToken || accessToken.length === 0) {
      return tenantId;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return tenantId;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      tenantId = token.tid;
    }
    catch {
    }

    return tenantId;
  },

  getUserNameFromAccessToken(accessToken: string): string {
    let userName: string = '';

    if (!accessToken || accessToken.length === 0) {
      return userName;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return userName;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      // if authenticated using certificate, there is no upn so use
      // app display name instead
      userName = token.upn || token.app_displayname;
    }
    catch {
    }

    return userName;
  },

  getUserIdFromAccessToken(accessToken: string): string {
    let userId: string = '';

    if (!accessToken || accessToken.length === 0) {
      return userId;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return userId;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      userId = token.oid;
    }
    catch {
    }

    return userId;
  },

  /**
   * Asserts the presence of a delegated access token.
   * @throws {CommandError} Will throw an error if the access token is not available.
   * @throws {CommandError} Will throw an error if the access token is an application-only access token.
   */
  assertDelegatedAccessToken(): void {
    const accessToken = auth?.connection?.accessTokens?.[auth.defaultResource]?.accessToken;
    if (!accessToken) {
      throw new CommandError('No access token found.');
    }

    if (this.isAppOnlyAccessToken(accessToken)) {
      throw new CommandError('This command does not support application-only permissions.');
    }
  }
};