import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export abstract class PackageRule extends JsonRule {
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
      const node = this.getAstNodeFromFile(project.packageJson, this.propertyName);
      this.addFindingWithPosition(findings, node);
    }
  }
}