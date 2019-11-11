import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../project-upgrade/model";
import { ExternalizeEntry } from "../model/ExternalizeEntry";

export class PnPJsRule extends BasicDependencyRule {
  private pnpModules = [
    {
      key: "@pnp/odata",
      globalName: "pnp.odata",
      globalDependencies: [
        "@pnp/common",
        "@pnp/logging",
        "tslib"
      ]
    },
    {
      key: "@pnp/common",
      globalName: "pnp.common",
    },
    {
      key: "@pnp/logging",
      globalName: "pnp.logging",
      globalDependencies: [
        "tslib"
      ]
    },
    {
      key: "@pnp/sp",
      globalName: "pnp.sp",
      globalDependencies: [
        "@pnp/logging",
        "@pnp/common",
        "@pnp/odata",
        "tslib"
      ]
    },
    {
      key: '@pnp/pnpjs',
      globalName: 'pnp'
    },
    {
      key: '@pnp/sp-taxonomy',
      globalName: 'pnp.sp-taxonomy',
      globalDependencies: [
        "@pnp/sp",
        "@pnp/common",
        "@pnp/sp-clientsvc"
      ]
    },
    {
      key: '@pnp/sp-clientsvc',
      globalName: 'pnp.sp-clientsvc',
      globalDependencies: [
        "@pnp/sp",
        "@pnp/logging",
        "@pnp/common",
        "@pnp/odata",
      ]
    }
  ];

  public visit(project: Project): Promise<ExternalizeEntry[]> {
    const findings = this.pnpModules
      .map(x => this.getModuleAndParents(project, x.key))
      .reduce((x, y) => [...x, ...y]);

    if (findings.filter(x => x.key !== '@pnp/pnpjs').length > 0) {
      findings.push({
        key: 'tslib',
        globalName: 'tslib',
        path: `https://unpkg.com/tslib@^1.10.0/tslib.js`
      });
    }

    return Promise.resolve(findings);
  }
  private getModuleAndParents(project: Project, moduleName: string): ExternalizeEntry[] {
    const result: ExternalizeEntry[] = [];
    const moduleConfiguration = this.pnpModules.find(x => x.key === moduleName);

    if (project.packageJson && moduleConfiguration) {
      const version: string | undefined = project.packageJson.dependencies[moduleName];

      if (version) {
        result.push({
          ...moduleConfiguration,
          path: `https://unpkg.com/${moduleConfiguration.key}@${version}/dist/${moduleName.replace('@pnp/', '')}.es5.umd${moduleName === '@pnp/common' || moduleName === ' @pnp/pnpjs' ? '.bundle' : ''}.min.js`,
        });
        moduleConfiguration.globalDependencies && moduleConfiguration.globalDependencies.forEach(dependency => {
          result.push(...this.getModuleAndParents(project, `@${dependency.replace('/', '.')}`));
        });
      }
    }
    
    return result;
  }
}