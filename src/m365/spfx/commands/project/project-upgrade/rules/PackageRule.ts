import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export abstract class PackageRule extends Rule {
  constructor(protected propertyName: string, protected add: boolean, protected propertyValue?: string) {
    super();
  }

  get title(): string {
    return this.propertyName;
  }

  get description(): string {
    return `${this.add ? 'Add' : 'Remove'} package.json property`;
  };

  get resolution(): string {
    return `{
  "${this.propertyName}": "${this.propertyValue}"
}`
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './package.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }
    
    if (!project.packageJson || 
        (this.add && !(project.packageJson as any)[this.propertyName]) ||
        (!this.add && (project.packageJson as any)[this.propertyName])) {
      this.addFinding(findings);
    }
  }
}