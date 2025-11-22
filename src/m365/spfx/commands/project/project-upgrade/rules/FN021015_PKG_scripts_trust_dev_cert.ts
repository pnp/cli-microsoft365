import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021015_PKG_scripts_trust_dev_cert extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021015';
  }

  get title(): string {
    return 'package.json scripts.trust-dev-cert';
  }

  get description(): string {
    return 'Add package.json scripts.trust-dev-cert property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "trust-dev-cert": "${this.script}"
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
      !project.packageJson.scripts['trust-dev-cert'] ||
      project.packageJson.scripts['trust-dev-cert'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.trust-dev-cert');
      this.addFindingWithPosition(findings, node);
    }
  }
}