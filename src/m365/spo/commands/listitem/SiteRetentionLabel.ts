export interface SiteRetentionLabel {
  AcceptMessagesOnlyFromSendersOrMembers: boolean;
  AccessType?: any;
  AllowAccessFromUnmanagedDevice?: any;
  AutoDelete: boolean;
  BlockDelete: boolean;
  BlockEdit: boolean;
  ComplianceFlags: number;
  ContainsSiteLabel: boolean;
  DisplayName: string;
  EncryptionRMSTemplateId?: any;
  HasRetentionAction: boolean;
  IsEventTag: boolean;
  MultiStageReviewerEmail?: any;
  NextStageComplianceTag?: any;
  Notes?: any;
  RequireSenderAuthenticationEnabled: boolean;
  ReviewerEmail?: any;
  SharingCapabilities?: any;
  SuperLock: boolean;
  TagDuration: number;
  TagId: string;
  TagName: string;
  TagRetentionBasedOn: string;
  UnlockedAsDefault: boolean;
} 