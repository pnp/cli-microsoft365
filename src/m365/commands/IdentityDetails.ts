export interface IdentityDetails {
  identityName: string;
  identityId: string;
  authType: string;
  appId: string;
  appTenant: string;
  accessTokens?: string;
  cloudType: string;
}