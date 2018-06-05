import { ConfigJson, CopyAssetsJson, DeployAzureStorageJson, Manifest, PackageJson, PackageSolutionJson, ServeJson, TsConfigJson, TsLintJson, WriteManifestsJson, YoRcJson } from ".";

export interface Project {
  path: string;

  configJson?: ConfigJson;
  copyAssetsJson?: CopyAssetsJson;
  deployAzureStorageJson?: DeployAzureStorageJson;
  manifests?: Manifest[];
  packageJson?: PackageJson;
  packageSolutionJson?: PackageSolutionJson;
  serveJson?: ServeJson;
  tsConfigJson?: TsConfigJson;
  tsLintJson?: TsLintJson;
  writeManifestsJson?: WriteManifestsJson;
  yoRcJson?: YoRcJson;
}