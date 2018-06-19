export interface PackageSolutionJson {
  $schema: string;
  solution?: {
    skipFeatureDeployment?: string | boolean;
  }
}