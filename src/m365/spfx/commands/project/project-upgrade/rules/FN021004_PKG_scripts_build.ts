import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021004_PKG_scripts_build extends JsonRule {
  private script: string;

  constructor(options: { script: string }) {
    super();
    this.script = options.script;
  }

  get id(): string {
    return 'FN021004';
  }

  get title(): string {
    return 'package.json scripts.build';
  }

  get description(): string {
    return 'Update package.json scripts.build property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "build": "${this.script}"
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
      !project.packageJson.scripts.build ||
      project.packageJson.scripts.build !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.build');
      this.addFindingWithPosition(findings, node);
    }
  }
}