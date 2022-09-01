import { JsonRule } from "../../JsonRule";
import { Project } from "../../project-model";
import { Finding } from "../../report-model";

export class FN021003_PKG_engines_node extends JsonRule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN021003';
  }

  get title(): string {
    return 'package.json engines.node';
  }

  get description(): string {
    return 'Update package.json engines.node property';
  }

  get resolution(): string {
    return `{
  "engines": {
    "node": "${this.version}"
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }

    if (!project.packageJson.engines ||
      !project.packageJson.engines.node ||
      project.packageJson.engines.node !== this.version) {
      const node = this.getAstNodeFromFile(project.packageJson, 'engines.node');
      this.addFindingWithPosition(findings, node);
    }
  }
}