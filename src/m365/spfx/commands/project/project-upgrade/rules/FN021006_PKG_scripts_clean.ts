import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021006_PKG_scripts_clean extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021006';
  }

  get title(): string {
    return 'package.json scripts.clean';
  }

  get description(): string {
    return 'Update package.json scripts.clean property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "clean": "${this.script}"
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

    if (!project.packageJson.scripts ||
      typeof project.packageJson.scripts !== 'object' ||
      !project.packageJson.scripts.clean ||
      project.packageJson.scripts.clean !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.clean');
      this.addFindingWithPosition(findings, node);
    }
  }
}