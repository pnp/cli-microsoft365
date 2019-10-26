import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";
import * as fs from 'fs';
import request from '../../../../../../request';


export class DynamicRule extends BasicDependencyRule {
  public async visit(project: Project): Promise<ExternalizeEntry[]> {
    if(project.packageJson) {
      const validPackageNames = Object.getOwnPropertyNames(project.packageJson.dependencies)
                            .filter(x => !this.restrictedNamespaces.map(y => x.indexOf(y) === -1).reduce((y, z) => y || z))
                            .filter(x => this.restrictedModules.indexOf(x) === -1);

      return (await Promise.all(validPackageNames.map((x) => this.getExternalEntryForPackage(x, project))))
              .filter(x => x !== undefined).map(x => x as ExternalizeEntry);
    } else {
      return [];
    }
  }
  private getExternalEntryForPackage = async (packageName: string, project: Project): Promise<ExternalizeEntry | undefined> => {
    const version = project.packageJson && project.packageJson.dependencies[packageName];
    const filePath = this.cleanFilePath(this.getFilePath(packageName));
    if(version && filePath) {
      let url = this.getFileUrl(packageName, version, filePath);
      let testResult = await this.testUrl(url);
      if (!url.endsWith('.min.js')) {
        const minUrl = url.replace('.js', '.min.js');
        const minResult = await this.testUrl(minUrl);
        if(minResult) {
          url = minUrl;
          testResult = true;
        }
      }
      if(testResult) {
        return {
          key: packageName,
          path: url,
        } as ExternalizeEntry;
      }
    }
    return undefined;
  }
  private getFileUrl = (packageName: string, version: string, filePath: string) => `https://unpkg.com/${packageName}@${version}/${filePath}`;
  private testUrl = async (url: string): Promise<boolean> => {
    try {
      await request.head({url: url, headers: {'x-anonymous': 'true'}});
      return true;
    } catch {
      return false;
    }
  }
  private getFilePath = (packageName: string): string | undefined => {
    const packageJsonFilePath = `node_modules/${packageName}/package.json`;
    try {
      const packageJson = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));
      if(packageJson.module) {
        return packageJson.module;
      } else if (packageJson.main) {
        return packageJson.main;
      } else {
        return undefined;
      }
    } catch { // file doesn't exist, giving up
      return undefined;
    }
  }
  private cleanFilePath = (filePath: string | undefined): string | undefined => {
    if(filePath) {
      return filePath.replace('./', '');
    } else {
      return undefined;
    }
  }
  private restrictedModules = ['react', 'react-dom'];
  private restrictedNamespaces = ['@types/', '@microsoft/'];
}