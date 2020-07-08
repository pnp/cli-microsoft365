export interface OrgAssets {
  Url: String;
  Libraries: OrgAssetsLibrary[];
}

export interface OrgAssetsLibrary{
  DisplayName: string;
  LibraryUrl: string;
  ListId: string;
  ThumbnailUrl?: string | null;
}

export interface OrgAssetsResponse {
  _ObjectType_: string;
  CentralAssetRepositoryLibraries?: any;
  OrgAssetsLibraries: OrgAssetsLibrariesResponse;
  SiteId: string;
  Url: Url;
  WebId: string;
}

export interface OrgAssetsLibrariesResponse {
  _ObjectType_: string;
  _Child_Items_: ChildItem[];
}

export interface ChildItem {
  _ObjectType_: string;
  DisplayName: string;
  FileType: string;
  LibraryUrl: Url;
  ListId: string;
  OrgAssetType: number;
  ThumbnailUrl: Url;
  UniqueId: string;
}

export interface Url {
  _ObjectType_: string;
  DecodedUrl: string;
}
