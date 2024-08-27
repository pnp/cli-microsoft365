import auth from '../Auth.js';
import { CommandError } from '../Command.js';

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

  getClaimsFromAccessToken(accessToken: string, ...claimNames: string[]): { [claimName: string]: string | number | undefined } | undefined {
    if (!accessToken || accessToken.length === 0) {
      return undefined;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return undefined;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);

      const claimsObject: { [claimName: string]: string | number | undefined } = claimNames.reduce((claimsObject: any, claimName: string) => {
        const claimValue = token[claimName];

        if (claimValue) {
          claimsObject[claimName] = token[claimName];
        }

        return claimsObject;
      }, {});

      return claimsObject;
    }
    catch {
    }

    return;
  },

  getTenantIdFromAccessToken(accessToken: string): string {
    const claims = this.getClaimsFromAccessToken(accessToken, 'tid');
    return claims?.tid as string || '';
  },

  getUserNameFromAccessToken(accessToken: string): string {
    const claims = this.getClaimsFromAccessToken(accessToken, 'upn', 'app_displayname');
    return claims?.upn as string || claims?.app_displayname as string || '';
  },

  getUserIdFromAccessToken(accessToken: string): string {
    const claims = this.getClaimsFromAccessToken(accessToken, 'oid');
    return claims?.oid as string || '';
  },

  getAppIdFromAccessToken(accessToken: string): string {
    const claims = this.getClaimsFromAccessToken(accessToken, 'appid');
    return claims?.appid as string || '';
  },

  getAudienceFromAccessToken(accessToken: string): string {
    const claims = this.getClaimsFromAccessToken(accessToken, 'aud');
    return claims?.aud as string || '';
  },

  getExpirationFromAccessToken(accessToken: string): Date | undefined {
    const claims = this.getClaimsFromAccessToken(accessToken, 'exp');

    if (!claims?.exp) {
      return undefined;
    }

    return new Date(claims.exp as number * 1000);
  },

  /**
   * Asserts the presence of a delegated access token.
   * @throws {CommandError} Will throw an error if the access token is not available.
   * @throws {CommandError} Will throw an error if the access token is an application-only access token.
   */
  assertDelegatedAccessToken(): void {
    const accessToken = auth?.connection?.accessTokens?.[Object.keys(auth.connection.accessTokens)[0]]?.accessToken;
    if (!accessToken) {
      throw new CommandError('No access token found.');
    }

    if (this.isAppOnlyAccessToken(accessToken)) {
      throw new CommandError('This command does not support application-only permissions.');
    }
  }
};