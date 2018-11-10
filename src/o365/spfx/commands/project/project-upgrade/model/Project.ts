import { ConfigJson, CopyAssetsJson, DeployAzureStorageJson, Manifest, PackageJson, PackageSolutionJson, ServeJson, TsConfigJson, TsLintJson, WriteManifestsJson, YoRcJson, GulpfileJs, VsCode, TsFile } from ".";

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
  tsFiles?: TsFile[];
  tsLintJson?: TsLintJson;
  tsLintJsonRoot?: TsLintJson;
  vsCode?: VsCode;
  writeManifestsJson?: WriteManifestsJson;
  yoRcJson?: YoRcJson;
}