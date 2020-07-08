export interface ListInstance {
  AllowContentTypes: boolean;
  BaseTemplate: number;
  BaseType: number;
  ContentTypesEnabled: boolean;
  CrawlNonDefaultViews: boolean;
  Created: string;
  CurrentChangeToken: Object;
  CustomActionElements: Object[];
  DefaultContentApprovalWorkflowId: string;
  DefaultItemOpenUseListSetting: boolean;
  Description: string;
  Direction: string;
  DocumentTemplateUrl: null;
  DraftVersionVisibility: number;
  EnableAttachments: boolean;
  EnableFolderCreation: boolean;
  EnableMinorVersions: boolean;
  EnableModeration: boolean;
  EnableVersioning: boolean;
  EntityTypeName: string;
  ExemptFromBlockDownloadOfNonViewableFiles: boolean;
  FileSavePostProcessingEnabled: boolean;
  ForceCheckout: boolean;
  HasExternalDataSource: boolean;
  Hidden: boolean;
  Id: string;
  ImagePath: Object;
  ImageUrl: string;
  IrmEnabled: boolean;
  IrmExpire: boolean;
  IrmReject: boolean;
  IsApplicationList: boolean;
  IsCatalog: boolean;
  IsPrivate: boolean;
  ItemCount: number;
  LastItemDeletedDate: string;
  LastItemModifiedDate: string;
  LastItemUserModifiedDate: string;
  ListExperienceOptions: number;
  ListItemEntityTypeFullName: string;
  MajorVersionLimit: number;
  MajorWithMinorVersionsLimit: number;
  MultipleDataList: boolean;
  NoCrawl: boolean;
  ParentWebPath: Object;
  ParentWebUrl: string;
  ParserDisabled: boolean;
  ServerTemplateCanCreateFolders: boolean;
  TemplateFeatureId: string;
  Title: string;
  RootFolder: RootFolder;
}

export interface RootFolder {
  ServerRelativeUrl: string;
}