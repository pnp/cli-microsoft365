import { ConfigJson, CopyAssetsJson, DeployAzureStorageJson, Gitignore, GulpfileJs, Manifest, PackageJson, PackageSolutionJson, ScssFile, ServeJson, TsConfigJson, TsFile, TsLintJson, VsCode, WriteManifestsJson, YoRcJson } from ".";
import { Npmignore } from "./Npmignore";

export interface Project {
  path: string;

  configJson?: ConfigJson;
  copyAssetsJson?: CopyAssetsJson;
  deployAzureStorageJson?: DeployAzureStorageJson;
  gitignore?: Gitignore;
  gulpfileJs?: GulpfileJs;
  manifests?: Manifest[];
  npmignore?: Npmignore;
  packageJson?: PackageJson;
  packageSolutionJson?: PackageSolutionJson;
  serveJson?: ServeJson;
  tsConfigJson?: TsConfigJson;
  tsFiles?: TsFile[];
  scssFiles?: ScssFile[];
  tsLintJson?: TsLintJson;
  tsLintJsonRoot?: TsLintJson;
  vsCode?: VsCode;
  writeManifestsJson?: WriteManifestsJson;
  yoRcJson?: YoRcJson;
}