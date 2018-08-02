import { ConfigJson, CopyAssetsJson, DeployAzureStorageJson, Manifest, PackageJson, PackageSolutionJson, ServeJson, TsConfigJson, TsLintJson, WriteManifestsJson, YoRcJson, GulpfileJs, VsCode } from ".";

export interface Project {
  path: string;

  configJson?: ConfigJson;
  copyAssetsJson?: CopyAssetsJson;
  deployAzureStorageJson?: DeployAzureStorageJson;
  gulpfileJs?: GulpfileJs;
  manifests?: Manifest[];
  packageJson?: PackageJson;
  packageSolutionJson?: PackageSolutionJson;
  serveJson?: ServeJson;
  tsConfigJson?: TsConfigJson;
  tsLintJson?: TsLintJson;
  vsCode?: VsCode;
  writeManifestsJson?: WriteManifestsJson;
  yoRcJson?: YoRcJson;
}