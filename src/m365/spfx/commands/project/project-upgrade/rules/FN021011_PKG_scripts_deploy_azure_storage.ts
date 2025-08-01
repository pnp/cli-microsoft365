import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021011_PKG_scripts_deploy_azure_storage extends JsonRule {
  constructor(private script: string) {
    super();
  }

  get id(): string {
    return 'FN021011';
  }

  get title(): string {
    return 'package.json scripts.deploy-azure-storage';
  }

  get description(): string {
    return 'Update package.json scripts.deploy-azure-storage property';
  }

  get resolution(): string {
    return `{
  "scripts": {
    "deploy-azure-storage": "${this.script}"
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
      !project.packageJson.scripts['deploy-azure-storage'] ||
      project.packageJson.scripts['deploy-azure-storage'] !== this.script) {
      const node = this.getAstNodeFromFile(project.packageJson, 'scripts.deploy-azure-storage');
      this.addFindingWithPosition(findings, node);
    }
  }
}