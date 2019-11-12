import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";
import * as fs from 'fs';
import request from '../../../../../../request';

export class DynamicRule extends BasicDependencyRule {
  private restrictedModules = ['react', 'react-dom'];
  private restrictedNamespaces = ['@types/', '@microsoft/'];

  public async visit(project: Project): Promise<ExternalizeEntry[]> {
    if (!project.packageJson) {
      return [];
    }

    const validPackageNames: string[] = Object.getOwnPropertyNames(project.packageJson.dependencies)
      .filter(x => this.restrictedNamespaces.map(y => x.indexOf(y) === -1).reduce((y, z) => y && z))
      .filter(x => this.restrictedModules.indexOf(x) === -1);

    return (await Promise.all(validPackageNames.map((x) => this.getExternalEntryForPackage(x, project))))
      .filter(x => x !== undefined)
      .map(x => x as ExternalizeEntry);
  }

  private async getExternalEntryForPackage(packageName: string, project: Project): Promise<ExternalizeEntry | undefined> {
    const version: string | undefined = project.packageJson && project.packageJson.dependencies[packageName];
    const filePath: string | undefined = this.cleanFilePath(this.getFilePath(packageName));

    if (!version || !filePath) {
      return undefined;
    }

    let url: string = this.getFileUrl(packageName, version, filePath);
    let testResult: boolean = await this.testUrl(url);

    if (!testResult) {
      return undefined;
    }

    if (!url.endsWith('.min.js')) {
      const minUrl: string = url.replace('.js', '.min.js');
      const minResult: boolean = await this.testUrl(minUrl);

      if (minResult) {
        url = minUrl;
        testResult = true;
      }
    }

    const moduleType = await this.getModuleType(url);

    return {
      key: packageName,
      path: url,
      globalName: moduleType === 'script' ? packageName : undefined,
    } as ExternalizeEntry;
  }

  private async getModuleType(url: string): Promise<'script' | 'module'> {
    try {
      const response = await request.post<string>({
        url: 'https://scriptcheck-weu-fn.azurewebsites.net/api/script-check',
        headers: { 'content-type': 'application/json', accept: 'application/json', 'x-anonymous': 'true' },
        body: JSON.stringify({ url: url }),
      });

      return JSON.parse(response).scriptType;
    }
    catch (error) {
      return 'module';
    }
  }

  private getFileUrl(packageName: string, version: string, filePath: string) {
    return `https://unpkg.com/${packageName}@${version}/${filePath}`;
  }

  private async testUrl(url: string): Promise<boolean> {
    try {
      await request.head({ url: url, headers: { 'x-anonymous': 'true' } });
      return true;
    }
    catch {
      return false;
    }
  }

  private getFilePath(packageName: string): string | undefined {
    const packageJsonFilePath: string = `node_modules/${packageName}/package.json`;

    try {
      const packageJson: { module?: any, main?: any } = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));

      if (packageJson.module) {
        return packageJson.module;
      }
      else if (packageJson.main) {
        return packageJson.main;
      }
      else {
        return undefined;
      }
    }
    catch { // file doesn't exist, giving up
      return undefined;
    }
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