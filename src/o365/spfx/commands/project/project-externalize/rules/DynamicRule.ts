import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../model";
import * as fs from 'fs';
import request from '../../../../../../request';
import { ExternalizeEntry } from "../";
import { VisitationResult } from '../VisitationResult';

type ModuleType = 'CommonJs' | 'UMD' | 'AMD' | 'ES2015' | 'non-module';

interface ScriptCheckApiResponse {
  scriptType: ModuleType;
  exports: string[];
}

export class DynamicRule extends BasicDependencyRule {
  private restrictedModules = ['react', 'react-dom', '@pnp/sp-clientsvc', '@pnp/sp-taxonomy'];
  private restrictedNamespaces = ['@types/', '@microsoft/'];
  private fileVariationSuffixes = ['.min', '.bundle', '-min', '.bundle.min'];

  public visit(project: Project): Promise<VisitationResult> {
    if (!project.packageJson) {
      return Promise.resolve({ entries: [], suggestions: [] });
    }

    const validPackageNames: string[] = Object.getOwnPropertyNames(project.packageJson.dependencies)
      .filter(x => this.restrictedNamespaces.map(y => x.indexOf(y) === -1).reduce((y, z) => y && z))
      .filter(x => this.restrictedModules.indexOf(x) === -1);

    return Promise
      .all(validPackageNames.map((x) => this.getExternalEntryForPackage(x, project)))
      .then((res: (ExternalizeEntry | undefined)[]) => {
        return {
          entries: res
            .filter(x => x !== undefined)
            .map(x => x as ExternalizeEntry),
          suggestions: []
        };
      });
  }

  private getExternalEntryForPackage(packageName: string, project: Project): Promise<ExternalizeEntry | undefined> {
    const version: string | undefined = project.packageJson && project.packageJson.dependencies[packageName];
    const filesPaths: string[] = this.getFilePath(packageName).map(x => this.cleanFilePath(x));

    if (!version || filesPaths.length === 0) {
      return Promise.resolve(undefined);
    }

    const filesPathsVariations = filesPaths
      .map(x => this.fileVariationSuffixes.map(y => x.indexOf(y) === -1 ? x.replace('.js', `${y}.js`) : x))
      .reduce((x, y) => [...x, ...y]);
    const pathsAndVariations = [...filesPaths, ...filesPathsVariations];
    return Promise.all(pathsAndVariations.map(x => this.getExternalEntryForFilePath(x, packageName, version)))
      .then((externalizeEntryCandidates) => {
        const dExternalizeEntryCandidates = externalizeEntryCandidates.filter(x => x !== undefined) as ExternalizeEntry[];
        const minifiedModule = dExternalizeEntryCandidates.find(x => !x.globalName && this.pathContainsMinifySuffix(x.path));
        const minifiedNonModule = dExternalizeEntryCandidates.find(x => x.globalName && this.pathContainsMinifySuffix(x.path));
        const nonMinifiedModule = dExternalizeEntryCandidates.find(x => x.globalName && !this.pathContainsMinifySuffix(x.path));
        const nonMinifiedNonModule = dExternalizeEntryCandidates.find(x => !x.globalName && !this.pathContainsMinifySuffix(x.path));
        const externalizeEntriesPrioritized = [minifiedModule, minifiedNonModule, nonMinifiedModule, nonMinifiedNonModule].filter(x => x !== undefined);
        return externalizeEntriesPrioritized.length > 0 ? externalizeEntriesPrioritized[0] : undefined;
      });
  }

  private pathContainsMinifySuffix(path: string): boolean {
    return this.fileVariationSuffixes
      .map(y => path.indexOf(y))
      .filter(y => y > -1).length > 0;
  }

  private getExternalEntryForFilePath(filePath: string, packageName: string, version: string): Promise<ExternalizeEntry | undefined> {
    let url: string = this.getFileUrl(packageName, version, filePath);

    return this
      .testUrl(url)
      .then((testResult) => {
        if (!testResult) {
          return Promise.resolve(undefined);
        }

        return this.getModuleType(url).then((moduleInfo) => {
          if (moduleInfo.scriptType === 'CommonJs') {
            return Promise.resolve(undefined); //browsers don't support those module types without an additional library
          } else if (moduleInfo.scriptType === 'ES2015' || moduleInfo.scriptType === 'AMD') {
            return {
              key: packageName,
              path: url,
            } as ExternalizeEntry;
          } else { //TODO for non-module and UMD we should technically add dependencies as well
            return {
              key: packageName,
              path: url,
              globalName: moduleInfo.exports && moduleInfo.exports.length > 0 ? moduleInfo.exports[0] : packageName, // examples where this is not good https://unpkg.com/@pnp/polyfill-ie11@^1.0.2/dist/index.js https://unpkg.com/moment-timezone@^0.5.27/builds/moment-timezone-with-data.js
            } as ExternalizeEntry;
          }
        });
      });
  }

  private getModuleType(url: string): Promise<ScriptCheckApiResponse> {
    return request
      .post<ScriptCheckApiResponse>({
        url: 'https://scriptcheck-weu-fn.azurewebsites.net/api/v2/script-check',
        headers: { 'content-type': 'application/json', accept: 'application/json', 'x-anonymous': 'true' },
        body: { url: url },
        json: true
      }).catch(() => {
        return { scriptType: 'non-module' as ModuleType, exports: [] };
      });
  }

  private getFileUrl(packageName: string, version: string, filePath: string) {
    return `https://unpkg.com/${packageName}@${version}/${filePath}`;
  }

  private testUrl(url: string): Promise<boolean> {
    return request
      .head({ url: url, headers: { 'x-anonymous': 'true' } })
      .then(() => true)
      .catch(() => false);
  }

  private getFilePath(packageName: string): string[] {
    const packageJsonFilePath: string = `node_modules/${packageName}/package.json`;
    let filesPaths: string[] = [];

    try {
      const packageJson: {
        module?: string,
        main?: string,
        es2015?: string,
        jspm?: {
          main?: string,
          files?: string[],
        }
        spm?: {
          main?: string,
        }
      } = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));
      if (packageJson.module) {
        filesPaths.push(packageJson.module);
      }
      if (packageJson.main) {
        filesPaths.push(packageJson.main);
      }
      if (packageJson.es2015) {
        filesPaths.push(packageJson.es2015);
      }
      if (packageJson.jspm) {
        if (packageJson.jspm.main) {
          filesPaths.push(packageJson.jspm.main);
        }
        if (packageJson.jspm.files) {
          filesPaths.push(...packageJson.jspm.files);
        }
      }
      if (packageJson.spm) {
        if (packageJson.spm.main) {
          filesPaths.push(packageJson.spm.main);
        }
      }
      filesPaths = filesPaths
        .filter((x, idx) => filesPaths.indexOf(x) === idx) //filtering duplicates
        .map(x => x.endsWith('.js') ? x : `${x}.js`); //some package managers omit the file extension
    }
    catch {
      // file doesn't exist, giving up
    }

    return filesPaths;
  }

  private cleanFilePath(filePath: string): string {
    return filePath.replace('./', '');
  }
}