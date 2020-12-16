import { lt, valid, validRange } from "semver";
import { Finding, Hash } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export abstract class DependencyRule extends JsonRule {
  constructor(protected packageName: string, protected packageVersion: string, protected isDevDep: boolean = false, protected isOptional: boolean = false, protected add: boolean = true) {
    super();
  }

  get title(): string {
    return this.packageName;
  }

  get description(): string {
    return `${(this.add ? 'Upgrade' : 'Remove')} SharePoint Framework ${(this.isDevDep ? 'dev ' : '')}dependency package ${this.packageName}`;
  };

  get resolution(): string {
    return this.add ?
      `${(this.isDevDep ? 'installDev' : 'install')} ${this.packageName}@${this.packageVersion}` :
      `${(this.isDevDep ? 'uninstallDev' : 'uninstall')} ${this.packageName}`;
  };

  get resolutionType(): string {
    return 'cmd';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './package.json';
  };

  customCondition(project: Project): boolean {
    return false;
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }

    const projectDependencies: Hash | undefined = this.isDevDep ? project.packageJson.devDependencies : project.packageJson.dependencies;
    const versionEntry: string | null = projectDependencies ? projectDependencies[this.packageName] : '';
    const packageVersion: string | null = valid(versionEntry);
    const versionRange: string | null = validRange(versionEntry);
    if (this.add) {
      let jsonProperty: string = this.isDevDep ? 'devDependencies' : 'dependencies';

      if (versionEntry) {
        jsonProperty += `.${this.packageName}`;

        if (packageVersion) {
          if (lt(packageVersion, this.packageVersion)) {
            const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
            this.addFindingWithPosition(findings, node);
          }
        }
        else {
          if (versionRange) {
            const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
            this.addFindingWithPosition(findings, node);
          }
        }
      }
      else {
        if (!this.isOptional || this.customCondition(project)) {
          const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
          this.addFindingWithCustomInfo(this.packageName, this.description.replace('Upgrade', 'Install'), [{
            file: this.file,
            resolution: this.resolution,
            position: this.getPositionFromNode(node)
          }], findings);
        }
      }
    }
    else {
      let jsonProperty: string = `${(this.isDevDep ? 'devDependencies' : 'dependencies')}.${this.packageName}`;

      if (versionEntry) {
        const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
        this.addFindingWithPosition(findings, node);
      }
    }
  }
}