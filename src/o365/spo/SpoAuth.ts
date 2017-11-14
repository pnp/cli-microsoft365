import Auth, { Service } from '../../Auth';

export class Site extends Service {
  tenantId: string;
  url: string;

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
}

export default new SpoAuth(new Site(), '9bc3ab49-b65d-410a-85ad-de819febfddc');