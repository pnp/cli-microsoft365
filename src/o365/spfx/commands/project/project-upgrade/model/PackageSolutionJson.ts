export interface PackageSolutionJson {
  $schema: string;
  solution: {
    includeClientSideAssets?: boolean;
  }
}