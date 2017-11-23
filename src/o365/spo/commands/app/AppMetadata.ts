export interface AppMetadata {
  ID: string;
  Title: string;
  InstalledVersion: string;
  Deployed: boolean;
  AppCatalogVersion: string;
  CanUpgrade: boolean;
  CurrentVersionDeployed: boolean;
  IsClientSideSolution: boolean;
}