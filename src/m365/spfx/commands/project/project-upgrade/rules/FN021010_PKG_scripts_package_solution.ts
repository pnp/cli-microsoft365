import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021010_PKG_scripts_package_solution extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021010';
  }

  get title(): string {
    return 'package.json scripts.package-solution';
  }

  get description(): string {
    return 'Update package.json scripts.package-solution property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "package-solution": "${this.script}"
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
      !project.packageJson.scripts['package-solution'] ||
      project.packageJson.scripts['package-solution'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.package-solution');
      this.addFindingWithPosition(findings, node);
    }
  }
}