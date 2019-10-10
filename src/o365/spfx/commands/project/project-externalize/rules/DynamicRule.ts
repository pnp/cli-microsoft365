import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";
import * as fs from 'fs';

export class DynamicRule extends BasicDependencyRule {
  public get ModuleName () {
    return '*';
  }  
  public visit(project: Project, findings: ExternalizeEntry[]): void {
    if(project.packageJson) {
      const alreadyExternalized = findings.map(x => x.key);
      const validPackageNames = Object.getOwnPropertyNames(project.packageJson.dependencies)
                            .filter(x => x.indexOf(this.typesPrefix) === -1)
                            .filter(x => this.restrictedModules.indexOf(x) === -1)
                            .filter(x => alreadyExternalized.indexOf(x) === -1);

      validPackageNames.forEach((packageName) => {
        const version = project.packageJson && project.packageJson.dependencies[packageName];
        const filePath = this.cleanFilePath(this.getFilePath(packageName));
        if(version && filePath) {
          findings.push({
            key: packageName,
            path: `https://unpkg.com/${packageName}@${version}/${filePath}`,
          });
        }
      });
    }
  }
  //TODO try injecting min in name
  //TODO head test on url
  private getFilePath = (packageName: string): string | undefined => {
    const packageJsonFilePath = `node_modules/${packageName}/package.json`;
    const packageJson = JSON.parse(fs.readFileSync(packageJsonFilePath, 'utf8'));
    if(packageJson.module) {
      return packageJson.module;
    } else if (packageJson.main) {
      return packageJson.main;
    } else {
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
  private restrictedModules = ['react', 'react-dom', '@microsoft/decorators', '@microsoft/sp-application-base', '@microsoft/sp-core-library', '@microsoft/sp-lodash-subset', '@microsoft/sp-office-ui-fabric-core', '@microsoft/sp-page-context', '@microsoft/sp-webpart-base'];
  private typesPrefix = '@types/';
}