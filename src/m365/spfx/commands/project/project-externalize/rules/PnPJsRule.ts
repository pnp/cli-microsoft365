import { BasicDependencyRule } from "./BasicDependencyRule";
import { Project } from "../../model";
import { ExternalizeEntry, FileEdit } from "../";
import { VisitationResult } from '../VisitationResult';

export class PnPJsRule extends BasicDependencyRule {
  private pnpModules = [
    {
      key: "@pnp/odata",
      globalName: "pnp.odata",
      globalDependencies: [
        "@pnp/common",
        "@pnp/logging",
        "tslib"
      ],
      shadowRequire: "require(\"@pnp/odata\");",
    },
    {
      key: "@pnp/common",
      globalName: "pnp.common",
      shadowRequire: "require(\"@pnp/common\");",
    },
    {
      key: "@pnp/logging",
      globalName: "pnp.logging",
      globalDependencies: [
        "tslib"
      ],
      shadowRequire: "require(\"@pnp/logging\");",
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
    }
  ];

  public visit(project: Project): Promise<VisitationResult> {
    const findings: ExternalizeEntry[] = this.pnpModules
      .map(x => this.getModuleAndParents(project, x.key))
      .reduce((x, y) => [...x, ...y]);
    const files: string[] = this.getEntryFilesList(project);
    const rawFileEdits: FileEdit[][] = this.pnpModules.filter(x => findings.find(y => y.key === x.key) !== undefined)
      .filter(x => x.shadowRequire !== undefined)
      .map(x => files.map(y => ({
        action: "add",
        path: y,
        targetValue: x.shadowRequire
      } as FileEdit)))
    const fileEdits: FileEdit[] = rawFileEdits.length > 0 ? rawFileEdits.reduce((x, y) => [...x, ...y]) : [];
    if (findings.filter(x => x.key && x.key !== '@pnp/pnpjs').length > 0) { // we're adding tslib only if we found other packages that are not the bundle which already contains tslib
      findings.push({
        key: 'tslib',
        globalName: 'tslib',
        path: `https://unpkg.com/tslib@^1.10.0/tslib.js`
      });
      fileEdits.push(...files.map(x => ({
        action: "add",
        path: x,
        targetValue: 'require(\"tslib\");'
      } as FileEdit)));
    }

    return Promise.resolve({ entries: findings, suggestions: fileEdits });
  }

  private getEntryFilesList(project: Project): string[] {
    return project && project.manifests ? project.manifests.map(x => x.path.replace('.manifest.json', '.ts')) : [];
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