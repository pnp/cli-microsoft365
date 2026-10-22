import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021008_PKG_scripts_eject_webpack extends JsonRule {
  private script: string;

  constructor(options: { script: string }) {
    super();
    this.script = options.script;
  }

  get id(): string {
    return 'FN021008';
  }

  get title(): string {
    return 'package.json scripts.eject-webpack';
  }

  get description(): string {
    return 'Update package.json scripts.eject-webpack property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "eject-webpack": "${this.script}"
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
      !project.packageJson.scripts['eject-webpack'] ||
      project.packageJson.scripts['eject-webpack'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.eject-webpack');
      this.addFindingWithPosition(findings, node);
    }
  }
}