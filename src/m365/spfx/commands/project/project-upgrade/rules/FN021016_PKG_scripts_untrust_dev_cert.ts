import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021016_PKG_scripts_untrust_dev_cert extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021016';
  }

  get title(): string {
    return 'package.json scripts.untrust-dev-cert';
  }

  get description(): string {
    return 'Add package.json scripts.untrust-dev-cert property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "untrust-dev-cert": "${this.script}"
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
      !project.packageJson.scripts['untrust-dev-cert'] ||
      project.packageJson.scripts['untrust-dev-cert'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.untrust-dev-cert');
      this.addFindingWithPosition(findings, node);
    }
  }
}