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
      catch {
        // Do nothing
      }
    }

    const npmignorePath: string = path.join(projectRootPath, '.npmignore');
    if (fs.existsSync(npmignorePath)) {
      try {
        project.npmignore = {
          source: fs.readFileSync(npmignorePath, 'utf-8')
        };
      }
      catch {
        // Do nothing
      }
    }

    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'config.json'), project, 'configJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'copy-assets.json'), project, 'copyAssetsJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'deploy-azure-storage.json'), project, 'deployAzureStorageJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'package.json'), project, 'packageJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'package-solution.json'), project, 'packageSolutionJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'serve.json'), project, 'serveJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'tsconfig.json'), project, 'tsConfigJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'tslint.json'), project, 'tsLintJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'tslint.json'), project, 'tsLintJsonRoot');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'write-manifests.json'), project, 'writeManifestsJson');
    this.readAndParseJsonFile(path.join(projectRootPath, '.yo-rc.json'), project, 'yoRcJson');
    this.readAndParseJsonFile(path.join(projectRootPath, 'config', 'sass.json'), project, 'sassJson');

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

    this.readAndParseJsonFile(path.join(projectRootPath, '.vscode', 'settings.json'), project, 'vsCode.settingsJson');
    this.readAndParseJsonFile(path.join(projectRootPath, '.vscode', 'extensions.json'), project, 'vsCode.extensionsJson');
    this.readAndParseJsonFile(path.join(projectRootPath, '.vscode', 'launch.json'), project, 'vsCode.launchJson');

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
      catch {
        // Do nothing
      }
    }

    const packageJsonPath: string = path.resolve(this.projectRootPath as string, 'package.json');
    try {
      const packageJson: any = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
      if (packageJson &&
        packageJson.dependencies &&
        packageJson.dependencies['@microsoft/sp-core-library']) {
        const coreLibVersion: string = packageJson.dependencies['@microsoft/sp-core-library'];
        return coreLibVersion.replace(/[^0-9.]/g, '');
      }
    }
    catch {
      // Do nothing
    }

    return undefined;
  }

  private readAndParseJsonFile(filePath: string, project: Project, keyPath: string): Project {
    if (fs.existsSync(filePath)) {
      try {
        const source = formatting.removeSingleLineComments(fs.readFileSync(filePath, 'utf-8'));
        const keys = keyPath.split('.') as (keyof Project)[];
        let current: any = project;

        for (let i = 0; i < keys.length - 1; i++) {
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
