import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";
import * as fs from 'fs';
import request from '../../../../../../request';

export class DynamicRule extends BasicDependencyRule {
  private restrictedModules = ['react', 'react-dom', '@pnp/sp-clientsvc', '@pnp/sp-taxonomy'];
  private restrictedNamespaces = ['@types/', '@microsoft/'];

  public visit(project: Project): Promise<ExternalizeEntry[]> {
    return new Promise<ExternalizeEntry[]>((resolve: (result: ExternalizeEntry[]) => void, reject: (err: any) => void): void => {
      if (!project.packageJson) {
        return resolve([]);
      }

      const validPackageNames: string[] = Object.getOwnPropertyNames(project.packageJson.dependencies)
        .filter(x => this.restrictedNamespaces.map(y => x.indexOf(y) === -1).reduce((y, z) => y && z))
        .filter(x => this.restrictedModules.indexOf(x) === -1);

      Promise
        .all(validPackageNames.map((x) => this.getExternalEntryForPackage(x, project)))
        .then((res: (ExternalizeEntry | undefined)[]): void => {
          resolve(res
            .filter(x => x !== undefined)
            .map(x => x as ExternalizeEntry));
        });
    });
  }

  private getExternalEntryForPackage(packageName: string, project: Project): Promise<ExternalizeEntry | undefined> {
    return new Promise<ExternalizeEntry | undefined>((resolve: (externalEntry: ExternalizeEntry | undefined) => void, reject: (err: any) => void): void => {
      const version: string | undefined = project.packageJson && project.packageJson.dependencies[packageName];
      const filePath: string | undefined = this.cleanFilePath(this.getFilePath(packageName));

      if (!version || !filePath) {
        return resolve(undefined);
      }

      let url: string = `https://unpkg.com/${packageName}@${version}/${filePath}`;
      let minUrl: string = url;

      this
        .testUrl(url)
        .then((): Promise<void> => {

          if (!url.endsWith('.min.js')) {
            minUrl = url.replace('.js', '.min.js');
            return this.testUrl(minUrl);
          }
          else {
            return Promise.resolve();
          }
        })
        .then((): Promise<'script' | 'module'> => {
          url = minUrl;

          return this.getModuleType(url);
        })
        .then((moduleType: 'script' | 'module'): void => {
          resolve({
            key: packageName,
            path: url,
            globalName: moduleType === 'script' ? packageName : undefined,
          } as ExternalizeEntry);
        }, () => {
          return resolve(undefined);
        });
    });
  }

  private getModuleType(url: string): Promise<'script' | 'module'> {
    return new Promise<'script' | 'module'>((resolve: (scriptType: 'script' | 'module') => void, reject: (err: any) => void): void => {
      request
        .post<{ scriptType: 'script' | 'module' }>({
          url: 'https://scriptcheck-weu-fn.azurewebsites.net/api/script-check',
          headers: { 'content-type': 'application/json', accept: 'application/json', 'x-anonymous': 'true' },
          body: { url: url },
          json: true
        })
        .then((res: { scriptType: 'script' | 'module' }): void => {
          resolve(res.scriptType);
        }, (): void => {
          resolve('module');
        });
    });
  }

  private testUrl(url: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: () => void): void => {
      request
        .head({ url: url, headers: { 'x-anonymous': 'true' } })
        .then(() => {
          return resolve();
        }, () => {
          return reject();
        });
    });
  }

  private getFilePath(packageName: string): string | undefined {
    let result: string | undefined = undefined;

    const packageJsonFilePath: string = `node_modules/${packageName}/package.json`;

    try {
      const packageJson: { module?: any, main?: any } = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));

      if (packageJson.module) {
        result = packageJson.module;
      }
      else if (packageJson.main) {
        result = packageJson.main;
      }
    } catch {
      // file doesn't exist, giving up
    }
    
    return result;
  }

  private cleanFilePath(filePath: string | undefined): string | undefined {
    if (filePath) {
      return filePath.replace('./', '');
    }
    else {
      return undefined;
    }
  }
}