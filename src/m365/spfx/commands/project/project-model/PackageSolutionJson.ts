import { JsonFile } from ".";

export interface PackageSolutionJson extends JsonFile {
  $schema: string;
  solution?: {
    developer?: PackageSolutionJsonDeveloper;
    features?: PackageSolutionJsonFeature[];
    includeClientSideAssets?: boolean;
    isDomainIsolated?: boolean;
    metadata?: PackageSolutionJsonMetadata;
    name?: string;
    skipFeatureDeployment?: boolean;
    version?: string;
  }
}

interface PackageSolutionJsonDeveloper {
  mpnId?: string;
  name?: string;
  privacyUrl?: string;
  termOfUseUrl?: string;
  websiteUrl?: string;
}

interface PackageSolutionJsonMetadata {
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

interface PackageSolutionJsonFeature {
  description?: string;
  id?: string;
  title?: string;
  version?: string;
}