import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN005002_CFG_DAS_workingDir extends JsonRule {
  private workingDir: string;
  constructor(options: { workingDir: string }) {
    super();
    this.workingDir = options.workingDir;
  }

  get id(): string {
    return 'FN005002';
  }

  get title(): string {
    return 'deploy-azure-storage.json workingDir';
  }

  get description(): string {
    return `Update deploy-azure-storage.json workingDir`;
  }

  get resolution(): string {
    return `{
  "workingDir": "${this.workingDir}"
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './config/deploy-azure-storage.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.deployAzureStorageJson) {
      return;
    }

    if (project.deployAzureStorageJson.workingDir !== this.workingDir) {
      const node = this.getAstNodeFromFile(project.deployAzureStorageJson, 'workingDir');
      this.addFindingWithPosition(findings, node);
    }
  }
}