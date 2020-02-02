import { Finding, Hash } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";
import { lt, valid, validRange } from "semver";

export abstract class DependencyRule extends Rule {
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
      if (versionEntry) {
        if (packageVersion) {
          if (lt(packageVersion, this.packageVersion)) {
            this.addFinding(findings);
          }
        }
        else {
          if (versionRange) {
            this.addFinding(findings);
          }
        }
      }
      else {
        if (!this.isOptional || this.customCondition(project)) {
          this.addFindingWithCustomInfo(this.packageName, this.description.replace('Upgrade', 'Install'), [{
            file: this.file,
            resolution: this.resolution
          }], findings);
        }
      }
    }
    else {
      if (versionEntry) {
        this.addFinding(findings);
      }
    }
  }
}