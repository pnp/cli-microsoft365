import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021007_PKG_scripts_deploy extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021007';
  }

  get title(): string {
    return 'package.json scripts.deploy';
  }

  get description(): string {
    return 'Update package.json scripts.deploy property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "deploy": "${this.script}"
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
      !project.packageJson.scripts.deploy ||
      project.packageJson.scripts.deploy !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.deploy');
      this.addFindingWithPosition(findings, node);
    }
  }
}