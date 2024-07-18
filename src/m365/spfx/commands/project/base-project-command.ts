import fs from 'fs';
import path from 'path';
import { formatting } from '../../../../utils/formatting.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import { Manifest, Project, ScssFile, TsFile } from "./project-model/index.js";
import { CommandError } from '../../../../Command.js';


export abstract class BaseProjectCommand extends AnonymousCommand {
  protected projectRootPath: string | null = null;

  protected getProject(projectRootPath: string): Project {
    const project: Project = {
      path: projectRootPath
    };

    const gitignorePath: string = path.join(projectRootPath, '.gitignore');
    if (fs.existsSync(gitignorePath)) {
      try {
        project.gitignore = {
          source: fs.readFileSync(gitignorePath, 'utf-8')
        };
      }
      catch { }
    }

    const npmignorePath: string = path.join(projectRootPath, '.npmignore');
    if (fs.existsSync(npmignorePath)) {
      try {
        project.npmignore = {
          source: fs.readFileSync(npmignorePath, 'utf-8')
        };
      }
      catch { }
    }

    const configJsonPath: string = path.join(projectRootPath, 'config', 'config.json');
    this.readAndParseJsonFile(configJsonPath, project, 'configJson');

    const copyAssetsJsonPath: string = path.join(projectRootPath, 'config', 'copy-assets.json');
    this.readAndParseJsonFile(copyAssetsJsonPath, project, 'copyAssetsJson');

    const deployAzureStorageJsonPath: string = path.join(projectRootPath, 'config', 'deploy-azure-storage.json');
    this.readAndParseJsonFile(deployAzureStorageJsonPath, project, 'deployAzureStorageJson');

    const packageJsonPath: string = path.join(projectRootPath, 'package.json');
    this.readAndParseJsonFile(packageJsonPath, project, 'packageJson');

    const packageSolutionJsonPath: string = path.join(projectRootPath, 'config', 'package-solution.json');
    this.readAndParseJsonFile(packageSolutionJsonPath, project, 'packageSolutionJson');

    const serveJsonPath: string = path.join(projectRootPath, 'config', 'serve.json');
    this.readAndParseJsonFile(serveJsonPath, project, 'serveJson');

    const tsConfigJsonPath: string = path.join(projectRootPath, 'tsconfig.json');
    this.readAndParseJsonFile(tsConfigJsonPath, project, 'tsConfigJson');

    const tsLintJsonPath: string = path.join(projectRootPath, 'config', 'tslint.json');
    this.readAndParseJsonFile(tsLintJsonPath, project, 'tsLintJson');

    const tsLintJsonRootPath: string = path.join(projectRootPath, 'tslint.json');
    this.readAndParseJsonFile(tsLintJsonRootPath, project, 'tsLintJsonRoot');

    const writeManifestJsonPath: string = path.join(projectRootPath, 'config', 'write-manifests.json');
    this.readAndParseJsonFile(writeManifestJsonPath, project, 'writeManifestsJson');

    const yoRcJsonPath: string = path.join(projectRootPath, '.yo-rc.json');
    this.readAndParseJsonFile(yoRcJsonPath, project, 'yoRcJson');

    const gulpfileJsPath: string = path.join(projectRootPath, 'gulpfile.js');
    if (fs.existsSync(gulpfileJsPath)) {
      project.gulpfileJs = {
        source: fs.readFileSync(gulpfileJsPath, 'utf-8')
      };
    }

    const esLintRcJsPath: string = path.join(projectRootPath, '.eslintrc.js');
    if (fs.existsSync(esLintRcJsPath)) {
      project.esLintRcJs = new TsFile(esLintRcJsPath);
    }

    project.vsCode = {};
    const vsCodeSettingsPath: string = path.join(projectRootPath, '.vscode', 'settings.json');
    this.readAndParseJsonFile(vsCodeSettingsPath, project, 'vsCode.settingsJson');

    const vsCodeExtensionsPath: string = path.join(projectRootPath, '.vscode', 'extensions.json');
    this.readAndParseJsonFile(vsCodeExtensionsPath, project, 'vsCode.extensionsJson');

    const vsCodeLaunchPath: string = path.join(projectRootPath, '.vscode', 'launch.json');
    this.readAndParseJsonFile(vsCodeLaunchPath, project, 'vsCode.launchJson');

    const srcFiles: string[] = fsUtil.readdirR(path.join(projectRootPath, 'src')) as string[];

    const manifestFiles = srcFiles.filter(f => f.endsWith('.manifest.json'));
    const manifests: Manifest[] = manifestFiles.map((f) => {
      const manifestStr = formatting.removeSingleLineComments(fs.readFileSync(f, 'utf-8'));
      const manifest: Manifest = formatting.parseJsonWithBom(manifestStr);
      manifest.path = f;
      manifest.source = manifestStr;
      return manifest;
    });
    project.manifests = manifests;

    const tsFiles: string[] = srcFiles.filter(f => f.endsWith('.ts') || f.endsWith('.tsx'));
    project.tsFiles = tsFiles.map(f => new TsFile(f));

    const scssFiles: string[] = srcFiles.filter(f => f.endsWith('.scss'));
    project.scssFiles = scssFiles.map(f => new ScssFile(f));

    return project;
  }

  protected getProjectRoot(folderPath: string): string | null {
    const packageJsonPath: string = path.resolve(folderPath, 'package.json');

    if (fs.existsSync(packageJsonPath)) {
      return folderPath;
    }
    else {
      const parentPath: string = path.resolve(folderPath, `..${path.sep}`);

      if (parentPath !== folderPath) {
        return this.getProjectRoot(parentPath);
      }
      else {
        return null;
      }
    }
  }

  protected getProjectVersion(): string | undefined {
    const yoRcPath: string = path.resolve(this.projectRootPath as string, '.yo-rc.json');

    if (fs.existsSync(yoRcPath)) {
      try {
        const yoRc: any = JSON.parse(fs.readFileSync(yoRcPath, 'utf-8'));
        if (yoRc && yoRc['@microsoft/generator-sharepoint']) {
          const version: string | undefined = yoRc['@microsoft/generator-sharepoint'].version;

          if (version) {
            switch (yoRc['@microsoft/generator-sharepoint'].environment) {
              case 'onprem19':
                return '1.4.1';
              case 'onprem':
                return '1.1.0';
              default:
                return version;
            }
          }
        }
      }
      catch { }
    }

    const packageJsonPath: string = path.resolve(this.projectRootPath as string, 'package.json');
    try {
      const packageJson: any = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
      if (packageJson &&
        packageJson.dependencies &&
        packageJson.dependencies['@microsoft/sp-core-library']) {
        const coreLibVersion: string = packageJson.dependencies['@microsoft/sp-core-library'];
        return coreLibVersion.replace(/[^0-9\.]/g, '');
      }
    }
    catch { }

    return undefined;
  }

  protected readAndParseJsonFile(filePath: string, project: Project, keyPath: string): Project {
    if (fs.existsSync(filePath)) {
      try {
        const source = formatting.removeSingleLineComments(fs.readFileSync(filePath, 'utf-8'));
        const keys = keyPath.split('.') as (keyof Project)[];
        let current: any = project;

        for (let i = 0; i < keys.length - 1; i++) {
          if (current[keys[i]] === undefined) {
            current[keys[i]] = {};
          }
          current = current[keys[i]];
        }

        const finalKey = keys[keys.length - 1];
        current[finalKey] = JSON.parse(source);
        if (typeof current[finalKey] === 'object' && current[finalKey] !== null) {
          current[finalKey].source = source;
        }
      }
      catch (error) {
        throw new CommandError(`The file ${filePath} is not a valid JSON file or is not utf-8 encoded. Error: ${error}`);
      }
    }
    return project;
  }
}
