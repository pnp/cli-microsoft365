import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN005002_CFG_DAS_workingDir extends JsonRule {
  constructor(private workingDir: string) {
    super();
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