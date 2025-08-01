import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021009_PKG_scripts_build_watch extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021009';
  }

  get title(): string {
    return 'package.json scripts.build-watch';
  }

  get description(): string {
    return 'Update package.json scripts.build-watch property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "build-watch": "${this.script}"
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
      !project.packageJson.scripts['build-watch'] ||
      project.packageJson.scripts['build-watch'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.build-watch');
      this.addFindingWithPosition(findings, node);
    }
  }
}