import fs from 'fs';
import type { AccessToken } from "../Auth.js";
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
      // Do nothing
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
      // Do nothing
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
      // Do nothing
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
      // Do nothing
    }

    return userId;
  },

  getDecodedAccessToken(accessToken: string): { header: any; payload: any } {
    const chunks = accessToken.split('.');
    const headerString = Buffer.from(chunks[0], 'base64').toString();
    const payloadString = Buffer.from(chunks[1], 'base64').toString();

    const header = JSON.parse(headerString);
    const payload = JSON.parse(payloadString);
    return { header, payload };
  },

  getScopesFromAccessToken(accessToken: string): string[] {
    let scopes: string[] = [];

    if (!accessToken || accessToken.length === 0) {
      return scopes;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return scopes;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      if (token.scp?.length > 0) {
        scopes = token.scp.split(' ');
      }
    }
    catch {
      // Do nothing
    }

    return scopes;
  },

  readAccessTokenFromFile(filePath: string): AccessToken {
    const contents = fs.readFileSync(filePath, 'utf8').trim();
    let accessTokenString: string | undefined;
    let expiresOn: Date | null = null;

    try {
      const parsed: any = JSON.parse(contents);
      if (typeof parsed === 'string') {
        accessTokenString = parsed;
      }
      else {
        accessTokenString = parsed.access_token || parsed.accessToken;
        if (parsed.expires_on) {
          expiresOn = new Date(parseInt(parsed.expires_on, 10) * 1000);
        }
        else if (parsed.expiresOn) {
          expiresOn = typeof parsed.expiresOn === 'number' ?
            new Date(parsed.expiresOn * 1000) :
            new Date(parsed.expiresOn);
        }
      }
    }
    catch {
      accessTokenString = contents;
    }

    if (!accessTokenString) {
      throw new CommandError('Token file does not contain a valid access token.');
    }

    const chunks = accessTokenString.split('.');
    if (chunks.length !== 3) {
      throw new CommandError('Token file does not contain a valid access token.');
    }

    if (!expiresOn) {
      try {
        const payloadString = Buffer.from(chunks[1], 'base64').toString();
        const payload: any = JSON.parse(payloadString);
        if (payload.exp) {
          expiresOn = new Date(payload.exp * 1000);
        }
      }
      catch {
        // Do nothing
      }
    }

    return {
      accessToken: accessTokenString,
      expiresOn
    };
  },

  /**
   * Asserts the presence of a delegated or application-only access token.
   * @throws {CommandError} Will throw an error if the access token is not available.
   * @throws {CommandError} Will throw an error if the access token type is not correct.
   */
  assertAccessTokenType(type: 'delegated' | 'application'): void {
    const accessToken = auth?.connection?.accessTokens?.[auth.defaultResource]?.accessToken;
    if (!accessToken) {
      throw new CommandError('No access token found.');
    }

    const isAppAccessToken = this.isAppOnlyAccessToken(accessToken);
    if (type === 'delegated' && isAppAccessToken) {
      throw new CommandError('This command requires delegated permissions.');
    }
    if (type === 'application' && !isAppAccessToken) {
      throw new CommandError('This command requires application-only permissions.');
    }
  }
};