export interface SitePermission {
  id: string;
  roles: string[];
  grantedToIdentities: SitePermissionIdentitySet[];
  roles: string[];
}

export interface SitePermissionIdentitySet {
  application: SitePermissionGrantIdentity;
}

export interface SitePermissionGrantIdentity {
  displayName: string;
  id: string;
}