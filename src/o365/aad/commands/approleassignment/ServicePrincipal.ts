import { AppRoleAssignment } from "./AppRoleAssignment";

export interface AppRole {
  allowedMemberTypes: string[];
  description: string;
  displayName: string;
  id: string;
  isEnabled: boolean;
  value: string;
}

export interface InformationalUrls {
  termsOfService?: any;
  support?: any;
  privacy?: any;
  marketing?: any;
}

export interface Oauth2Permissions {
  adminConsentDescription: string;
  adminConsentDisplayName: string;
  id: string;
  isEnabled: boolean;
  type: string;
  userConsentDescription: string;
  userConsentDisplayName: string;
  value: string;
}

export interface ServicePrincipal {
  appRoleAssignments: AppRoleAssignment[];
  objectType: string;
  objectId: string;
  deletionTimestamp?: any;
  accountEnabled: boolean;
  addIns: any[];
  alternativeNames: any[];
  appDisplayName: string;
  appId: string;
  applicationTemplateId?: any;
  appOwnerTenantId: string;
  appRoleAssignmentRequired: boolean;
  appRoles: AppRole[];
  displayName: string;
  errorUrl?: any;
  homepage?: any;
  informationalUrls: InformationalUrls;
  keyCredentials: any[];
  logoutUrl?: any;
  notificationEmailAddresses: any[];
  oauth2Permissions: Oauth2Permissions[];
  passwordCredentials: any[];
  preferredSingleSignOnMode?: any;
  preferredTokenSigningKeyEndDateTime?: any;
  preferredTokenSigningKeyThumbprint?: any;
  publisherName: string;
  replyUrls: any[];
  samlMetadataUrl?: any;
  samlSingleSignOnSettings?: any;
  servicePrincipalNames: string[];
  servicePrincipalType: string;
  signInAudience: string;
  tags: any[];
  tokenEncryptionKeyId?: any;
}
