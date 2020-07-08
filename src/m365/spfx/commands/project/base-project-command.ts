import Command from "../../../../Command";
import { TsFile, Manifest, Project, ScssFile } from "./model";
import { Utils } from './project-upgrade/';
import * as path from 'path';
import * as fs from 'fs';

export abstract class BaseProjectCommand extends Command {
  protected projectRootPath: string | null = null;

  protected getProject(projectRootPath: string): Project {
    const project: Project = {
      path: projectRootPath
    };

    const configJsonPath: string = path.join(projectRootPath, 'config/config.json');
    if (fs.existsSync(configJsonPath)) {
      try {
        project.configJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(configJsonPath, 'utf-8')));
      }
      catch { }
    }

    const copyAssetsJsonPath: string = path.join(projectRootPath, 'config/copy-assets.json');
    if (fs.existsSync(copyAssetsJsonPath)) {
      try {
        project.copyAssetsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(copyAssetsJsonPath, 'utf-8')));
      }
      catch { }
    }

    const deployAzureStorageJsonPath: string = path.join(projectRootPath, 'config/deploy-azure-storage.json');
    if (fs.existsSync(deployAzureStorageJsonPath)) {
      try {
        project.deployAzureStorageJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(deployAzureStorageJsonPath, 'utf-8')));
      }
      catch { }
    }

    const packageJsonPath: string = path.join(projectRootPath, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      try {
        project.packageJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(packageJsonPath, 'utf-8')));
      }
      catch { }
    }

    const packageSolutionJsonPath: string = path.join(projectRootPath, 'config/package-solution.json');
    if (fs.existsSync(packageSolutionJsonPath)) {
      try {
        project.packageSolutionJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(packageSolutionJsonPath, 'utf-8')));
      }
      catch { }
    }

    const serveJsonPath: string = path.join(projectRootPath, 'config/serve.json');
    if (fs.existsSync(serveJsonPath)) {
      try {
        project.serveJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(serveJsonPath, 'utf-8')));
      }
      catch { }
    }

    const tsConfigJsonPath: string = path.join(projectRootPath, 'tsconfig.json');
    if (fs.existsSync(tsConfigJsonPath)) {
      try {
        project.tsConfigJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(tsConfigJsonPath, 'utf-8')));
      }
      catch { }
    }

    const tsLintJsonPath: string = path.join(projectRootPath, 'config/tslint.json');
    if (fs.existsSync(tsLintJsonPath)) {
      try {
        project.tsLintJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(tsLintJsonPath, 'utf-8')));
      }
      catch { }
    }

    const tsLintJsonRootPath: string = path.join(projectRootPath, 'tslint.json');
    if (fs.existsSync(tsLintJsonRootPath)) {
      try {
        project.tsLintJsonRoot = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(tsLintJsonRootPath, 'utf-8')));
      }
      catch { }
    }

    const writeManifestJsonPath: string = path.join(projectRootPath, 'config/write-manifests.json');
    if (fs.existsSync(writeManifestJsonPath)) {
      try {
        project.writeManifestsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(writeManifestJsonPath, 'utf-8')));
      }
      catch { }
    }

    const yoRcJsonPath: string = path.join(projectRootPath, '.yo-rc.json');
    if (fs.existsSync(yoRcJsonPath)) {
      try {
        project.yoRcJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(yoRcJsonPath, 'utf-8')));
      }
      catch { }
    }

    const gulpfileJsPath: string = path.join(projectRootPath, 'gulpfile.js');
    if (fs.existsSync(gulpfileJsPath)) {
      project.gulpfileJs = {
        src: fs.readFileSync(gulpfileJsPath, 'utf-8')
      };
    }

    project.vsCode = {};
    const vsCodeSettingsPath: string = path.join(projectRootPath, '.vscode', 'settings.json');
    if (fs.existsSync(vsCodeSettingsPath)) {
      try {
        project.vsCode.settingsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeSettingsPath, 'utf-8')));
      }
      catch { }
    }

    const vsCodeExtensionsPath: string = path.join(projectRootPath, '.vscode', 'extensions.json');
    if (fs.existsSync(vsCodeExtensionsPath)) {
      try {
        project.vsCode.extensionsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeExtensionsPath, 'utf-8')));
      }
      catch { }
    }

    const vsCodeLaunchPath: string = path.join(projectRootPath, '.vscode', 'launch.json');
    if (fs.existsSync(vsCodeLaunchPath)) {
      try {
        project.vsCode.launchJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeLaunchPath, 'utf-8')));
      }
      catch { }
    }

    const srcFiles: string[] = Utils.getAllFiles(path.join(projectRootPath, 'src')) as string[];

    const manifestFiles = srcFiles.filter(f => f.endsWith('.manifest.json'));
    const manifests: Manifest[] = manifestFiles.map((f) => {
      const manifestStr = Utils.removeSingleLineComments(fs.readFileSync(f, 'utf-8'));
      const manifest: Manifest = JSON.parse(manifestStr);
      manifest.path = f;
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
}