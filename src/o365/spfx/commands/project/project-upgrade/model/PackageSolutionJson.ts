export interface PackageSolutionJson {
  $schema: string;
  solution?: {
    includeClientSideAssets?: boolean;
    isDomainIsolated?: boolean;
    skipFeatureDeployment?: boolean;
  }
}