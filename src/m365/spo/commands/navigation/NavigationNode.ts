export interface NavigationNode {
  Id: number;
  IsDocLib: boolean;
  IsExternal: boolean;
  IsVisible: boolean;
  ListTemplateType: number;
  Title: string;
  Url: string;
  AudienceIds: string[];
}

export interface MenuStateNode {
  AudienceIds: string[];
  CurrentLCID: number;
  CustomProperties: any[];
  FriendlyUrlSegment: string;
  IsDeleted: boolean;
  IsHidden: boolean;
  IsTitleForExistingLanguage: boolean;
  Key: string;
  Nodes: MenuStateNode[];
  NodeType: number;
  OpenInNewWindow?: boolean | null;
  SimpleUrl: string;
  Title: string;
  Translations: any[];
}

export interface MenuState {
  AudienceIds: string[];
  FriendlyUrlPrefix: string;
  IsAudienceTargetEnabledForGlobalNav: boolean;
  Nodes: MenuStateNode[];
  SimpleUrl: string;
  SPSitePrefix: string;
  SPWebPrefix: string;
  StartingNodeKey: string;
  StartingNodeTitle: string;
  Version: string;
}