import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021014_PKG_scripts_test_only extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021014';
  }

  get title(): string {
    return 'package.json scripts.test-only';
  }

  get description(): string {
    return 'Add package.json scripts.test-only property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "test-only": "${this.script}"
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
      !project.packageJson.scripts['test-only'] ||
      project.packageJson.scripts['test-only'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test-only');
      this.addFindingWithPosition(findings, node);
    }
  }
}