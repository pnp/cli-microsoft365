export interface MenuStateNode {
  AudienceIds: string[];
  CurrentLCID: number;
  CustomProperties: any[];
  FriendlyUrlSegment: string;
  IsDeleted: boolean;
  IsHidden: boolean;
  IsTitleForExistingLanguage: boolean;
  Key: string | null;
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