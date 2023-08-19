export interface OrgAssets {
  Url: string;
  Libraries: OrgAssetsLibrary[];
}

interface OrgAssetsLibrary {
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

interface OrgAssetsLibrariesResponse {
  _ObjectType_: string;
  _Child_Items_: ChildItem[];
}

interface ChildItem {
  _ObjectType_: string;
  DisplayName: string;
  FileType: string;
  LibraryUrl: Url;
  ListId: string;
  OrgAssetType: number;
  ThumbnailUrl: Url;
  UniqueId: string;
}

interface Url {
  _ObjectType_: string;
  DecodedUrl: string;
}
