export interface AdminUserResult {
  email: string;
  loginName: string;
  name: string;
  userPrincipalName: string;
}

export interface AdminResult {
  value: AdminUserResult[];
}

export interface SiteUserResult {
  Email: string;
  Id: number;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
}

export interface SiteResult {
  value: SiteUserResult[];
}

export interface AdminCommandResultItem {
  Id?: number;
  Email: string;
  IsPrimaryAdmin: boolean;
  LoginName: string;
  Title: string;
  PrincipalType?: number;
  PrincipalTypeString?: string;
}

export interface IGraphUser {
  userPrincipalName: string;
}

export interface ISiteOwner {
  LoginName: string;
}

export interface ISPSite {
  Id: string;
}

export interface ISiteUser {
  Id: number;
}
