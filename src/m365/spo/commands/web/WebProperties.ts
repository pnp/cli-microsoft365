export interface WebProperties {
  AllowRssFeeds: boolean;
  AlternateCssUrl: string;
  AppInstanceId: string;
  AssociatedMemberGroup: AssociatedGroupProperties;
  AssociatedOwnerGroup: AssociatedGroupProperties;
  AssociatedVisitorGroup: AssociatedGroupProperties;
  Configuration: number;
  Created: string;
  CurrentChangeToken: CurrentChangeToken;
  CustomMasterUrl: string;
  Description: string;
  DesignPackageId: string;
  DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: boolean;
  EnableMinimalDownload: boolean;
  HorizontalQuickLaunch: boolean;
  Id: string;
  IsMultilingual: boolean;
  Language: number;
  LastItemModifiedDate: string;
  LastItemUserModifiedDate: string;
  MasterUrl: string;
  NoCrawl: boolean;
  OverwriteTranslationsOnChange: boolean;
  ResourcePath: ResourcePath;
  QuickLaunchEnabled: boolean;
  RecycleBinEnabled: boolean;
  ServerRelativeUrl: string;
  SiteLogoUrl: string;
  SyndicationEnabled: boolean;
  Title: string;
  TreeViewEnabled: boolean;
  UIVersion: number;
  UIVersionConfigurationEnabled: boolean;
  Url: string;
  WebTemplate: string;
}

export interface CurrentChangeToken {
  StringValue: string;
}

export interface ResourcePath {
  DecodedUrl: string;
}
export interface AssociatedGroupProperties {
  Id: number;
  IsHiddenInUI: boolean;
  LoginName: string;
  Title: string;
  PrincipalType: number;
  AllowMembersEditMembership: boolean;
  AllowRequestToJoinLeave: boolean;
  AutoAcceptRequestToJoinLeave: boolean;
  Description: string;
  OnlyAllowMembersViewMembership: boolean;
  OwnerTitle: string;
  RequestToJoinLeaveEmailSetting: string;
}