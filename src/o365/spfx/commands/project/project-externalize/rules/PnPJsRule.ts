import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

export class PnPJsRule extends BasicDependencyRule {
  public get ModuleName () {
    return '@pnp/pnpjs';
  }  
  public visit(project: Project, findings: ExternalizeEntry[]): void {
    if(project.packageJson) {
      const version = project.packageJson.dependencies[this.ModuleName];
      if(version) {
        findings.push({
          key: this.ModuleName,
          path: `https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/${version}/pnpjs.es5.umd.bundle.min.js`,
          globalName: 'pnp'
        });
      }
    }
  }
}