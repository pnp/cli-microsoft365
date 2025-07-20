import { ConfigJson, CopyAssetsJson, DeployAzureStorageJson, EsLintRcJs, Gitignore, GulpfileJs, Manifest, Npmignore, PackageJson, PackageSolutionJson, SassJson, ScssFile, ServeJson, TsConfigJson, TsFile, TsLintJson, VsCode, WriteManifestsJson, YoRcJson } from '.';

export interface Project {
  path: string;
  version?: string;

  configJson?: ConfigJson;
  copyAssetsJson?: CopyAssetsJson;
  deployAzureStorageJson?: DeployAzureStorageJson;
  esLintRcJs?: EsLintRcJs;
  gitignore?: Gitignore;
  gulpfileJs?: GulpfileJs;
  manifests?: Manifest[];
  npmignore?: Npmignore;
  packageJson?: PackageJson;
  packageSolutionJson?: PackageSolutionJson;
  serveJson?: ServeJson;
  tsConfigJson?: TsConfigJson;
  tsFiles?: TsFile[];
  sassJson?: SassJson;
  scssFiles?: ScssFile[];
  tsLintJson?: TsLintJson;
  tsLintJsonRoot?: TsLintJson;
  vsCode?: VsCode;
  writeManifestsJson?: WriteManifestsJson;
  yoRcJson?: YoRcJson;
}
