export interface SiteProperties {
  Status: string;
  Title: string;
  Url: string;
}

export interface RecycleBinItemProperties{
  AuthorEmail: string;
  AuthorName: string;
  DeletedByEmail: string;
  DeletedByName: string;
  DeletedDate: Date;
  DeletedDateLocalFormatted: string;
  DirName: string;
  DirNamePath: any;
  Id: string;
  ItemState: number;
  ItemType: number;
  LeafName: string;
  LeafNamePath: any;
  Size: string;
  Title: string;
}

export interface RecycleBinItemPropertiesCollection {
  value: RecycleBinItemProperties[];
} 