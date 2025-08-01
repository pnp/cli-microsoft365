import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021005_PKG_scripts_test extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021005';
  }

  get title(): string {
    return 'package.json scripts.test';
  }

  get description(): string {
    return 'Update package.json scripts.test property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "test": "${this.script}"
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
      !project.packageJson.scripts.test ||
      project.packageJson.scripts.test !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test');
      this.addFindingWithPosition(findings, node);
    }
  }
}