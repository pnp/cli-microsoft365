import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

export class PnPJsRule extends BasicDependencyRule {
  public get ModuleName () {
    return '@pnp/pnpjs';
  }  
  public visit(project: Project): Promise<ExternalizeEntry[]> {
    if(project.packageJson) {
      const version = project.packageJson.dependencies[this.ModuleName];
      if(version) {
        return Promise.resolve([{
          key: this.ModuleName,
          path: `https://unpkg.com/${this.ModuleName}@${version}/dist/pnpjs.es5.umd.bundle.min.js`,
          globalName: 'pnp'
        }]);
      }
    }
    return Promise.resolve([]);
  }
}