import { JsonFile } from ".";

export interface PackageSolutionJson extends JsonFile {
  $schema: string;
  solution?: {
    developer?: PackageSolutionJsonDeveloper;
    features?: PackageSolutionJsonFeature[];
    includeClientSideAssets?: boolean;
    isDomainIsolated?: boolean;
    metadata?: PackageSolutionJsonMetadata;
    skipFeatureDeployment?: boolean;
    version?: string;
  }
}

export interface PackageSolutionJsonDeveloper {
  mpnId?: string;
  name?: string;
  privacyUrl?: string;
  termOfUseUrl?: string;
  websiteUrl?: string;
}

export interface PackageSolutionJsonMetadata {
  categories?: string[];
  longDescription?: {
    default?: string;
  };
  screenshotPaths?: string[];
  shortDescription?: {
    default?: string;
  };
  videoUrl?: string;
}

export interface PackageSolutionJsonFeature {
  description?: string;
  id?: string;
  title?: string;
  version?: string;
}