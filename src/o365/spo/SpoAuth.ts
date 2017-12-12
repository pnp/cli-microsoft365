import Auth, { Service, Logger } from '../../Auth';
import { CommandError } from '../../Command';

interface Hash {
  [indexer: string]: Token;
}

interface Token {
  accessToken: string;
  expiresAt: number;
}

export class Site extends Service {
  tenantId: string;
  url: string;
  accessTokens: Hash = {};

  public disconnect(): void {
    super.disconnect();
    this.tenantId = '';
    this.url = '';
  }

  public isTenantAdminSite(): boolean {
    return this.url !== null &&
      this.url !== undefined &&
      this.url.indexOf('-admin.sharepoint.com') > -1;
  }
}

class SpoAuth extends Auth {
  private SERVICE: string = 'SPO';

  constructor(public site: Site, appId: string) {
    super(site, appId);
  }

  public restoreAuth(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .getServiceConnectionInfo<Site>(this.SERVICE)
        .then((site: Site): void => {
          const s: Site = new Site();
          for (let k in site) {
            (s as any)[k] = (site as any)[k];
          }
          this.site = s;
          this.service = s;
          resolve();
        }, (error: any): void => {
          resolve();
        });
    });
  }

  public ensureAccessToken(resource: string, stdout: Logger, debug: boolean = false): Promise<string> {
    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const now: number = new Date().getTime() / 1000;
      const token: Token | undefined = this.site.accessTokens[resource];
      if (token && token.expiresAt > now) {
        resolve(this.site.accessTokens[resource].accessToken);
        return;
      }

      super
        .ensureAccessToken(resource, stdout, debug)
        .then((accessToken: string): void => {
          this.site.accessTokens[resource] = {
            accessToken: accessToken,
            expiresAt: new Date().getTime() / 1000
          };
          this
            .setServiceConnectionInfo(this.SERVICE, this.site)
            .then((): void => {
              resolve(accessToken);
            }, (error: any): void => {
              if (debug) {
                stdout.log(new CommandError(error));
              }
              resolve(accessToken);
            });
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public getAccessToken(resource: string, refreshToken: string, stdout: Logger, debug: boolean = false): Promise<string> {
    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const now: number = new Date().getTime() / 1000;
      const token: Token | undefined = this.site.accessTokens[resource];
      if (token && token.expiresAt > now) {
        resolve(this.site.accessTokens[resource].accessToken);
        return;
      }

      super
        .getAccessToken(resource, refreshToken, stdout, debug)
        .then((accessToken: string): void => {
          this.site.accessTokens[resource] = {
            accessToken: accessToken,
            expiresAt: new Date().getTime() / 1000
          };
          this
            .setServiceConnectionInfo(this.SERVICE, this.site)
            .then((): void => {
              resolve(accessToken);
            }, (error: any): void => {
              if (debug) {
                stdout.log(new CommandError(error));
              }
              resolve(accessToken);
            });
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public storeSiteConnectionInfo(): Promise<void> {
    return this.setServiceConnectionInfo(this.SERVICE, this.site);
  }

  public clearSiteConnectionInfo(): Promise<void> {
    return this.clearServiceConnectionInfo(this.SERVICE);
  }
}

export default new SpoAuth(new Site(), '9bc3ab49-b65d-410a-85ad-de819febfddc');