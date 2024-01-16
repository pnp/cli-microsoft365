export interface IdentityDetails {
  connectionName: string;
  connectedAs: string;
  authType: string;
  appId: string;
  appTenant: string;
  accessTokens?: string;
  cloudType: string;
}