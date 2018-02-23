export interface WebProperties {
  Id: string;
  Title: string;
  RootFolder: RootFolder;
}

export interface RootFolder {
  ServerRelativeUrl: string;
}