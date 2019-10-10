import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

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
        if(version) {
          findings.push({
            key: packageName,
            path: `https://unpkg.com/${packageName}@${version}/dist/pnpjs.es5.umd.bundle.min.js`,
          });
        }
      });
    }
  }//TODO get filename from package.json first module then main
  //TODO try injecting min in name
  //TODO head test on url
  private restrictedModules = ['react', 'react-dom', '@microsoft/decorators', '@microsoft/sp-application-base', '@microsoft/sp-core-library', '@microsoft/sp-lodash-subset', '@microsoft/sp-office-ui-fabric-core', '@microsoft/sp-page-context', '@microsoft/sp-webpart-base'];
  private typesPrefix = '@types/'
}