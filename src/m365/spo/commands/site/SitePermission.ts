export interface SitePermission {
  id: string;
  grantedToIdentities: SitePermissionIdentitySet[];
}

export interface SitePermissionIdentitySet {
  application: SitePermissionGrantIdentity;
}

export interface SitePermissionGrantIdentity {
  displayName: string;
  id: string;
}