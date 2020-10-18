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
  logoUrl?: null,
  marketingUrl?: null,
  privacyStatementUrl?: null,
  supportUrl?: null,
  termsOfServiceUrl?: null
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
  id: string;
  deletedDateTime?: any;
  accountEnabled: boolean;
  addIns: any[];
  alternativeNames: any[];
  appDisplayName: string;
  appId: string;
  applicationTemplateId?: any;
  appOwnerOrganizationId: string;
  appRoleAssignmentRequired: boolean;
  appRoles: AppRole[];
  displayName: string;
  errorUrl?: any;
  homepage?: any;
  info: InformationalUrls;
  keyCredentials: any[];
  logoutUrl?: any;
  notificationEmailAddresses: any[];
  oauth2PermissionScopes: Oauth2Permissions[];
  passwordCredentials: any[];
  preferredSingleSignOnMode?: any;
  preferredTokenSigningKeyEndDateTime?: any;
  preferredTokenSigningKeyThumbprint?: any;
  verifiedPublisher?: any;
  replyUrls: any[];
  samlMetadataUrl?: any;
  samlSingleSignOnSettings?: any;
  servicePrincipalNames: string[];
  servicePrincipalType: string;
  signInAudience: string;
  tags: any[];
  tokenEncryptionKeyId?: any;
}
