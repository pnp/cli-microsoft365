import { Auth, Logger, Service } from "../../Auth";
import { CommandError } from '../../Command';
import config from "../../config";

interface Hash {
  [indexer: string]: Token;
}

interface Token {
  accessToken: string;
  expiresOn: string;
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
  constructor(public site: Site, appId: string) {
    super(site, appId);
  }

  protected serviceId(): string {
    return 'SPO';
  }

  public restoreAuth(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .getServiceConnectionInfo<Site>(this.serviceId())
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
      const now: Date = new Date();
      const token: Token | undefined = this.site.accessTokens[resource];
      if (token) {
        const tokenExpiresOn: Date = new Date(token.expiresOn);
        if (tokenExpiresOn > now) {
          resolve(this.site.accessTokens[resource].accessToken);
          return;
        }
      }

      super
        .ensureAccessToken(resource, stdout, debug)
        .then((accessToken: string): void => {
          this.site.accessTokens[resource] = {
            accessToken: accessToken,
            expiresOn: new Date().toISOString()
          };
          this
            .setServiceConnectionInfo(this.serviceId(), this.site)
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
      const now: Date = new Date();
      const token: Token | undefined = this.site.accessTokens[resource];
      if (token) {
        const tokenExpiresOn: Date = new Date(token.expiresOn);
        if (tokenExpiresOn > now) {
          resolve(this.site.accessTokens[resource].accessToken);
          return;
        }
      }

      super
        .getAccessToken(resource, refreshToken, stdout, debug)
        .then((accessToken: string): void => {
          this.site.accessTokens[resource] = {
            accessToken: accessToken,
            expiresOn: new Date().toISOString()
          };
          this
            .setServiceConnectionInfo(this.serviceId(), this.site)
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
    return this.setServiceConnectionInfo(this.serviceId(), this.site);
  }

  public clearSiteConnectionInfo(): Promise<void> {
    return this.clearServiceConnectionInfo(this.serviceId());
  }
}

export default new SpoAuth(new Site(), config.aadSpoAppId);