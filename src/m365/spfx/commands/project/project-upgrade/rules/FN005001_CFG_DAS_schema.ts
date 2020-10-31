import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN005001_CFG_DAS_schema extends Rule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN005001';
  }

  get title(): string {
    return 'deploy-azure-storage.json schema';
  }

  get description(): string {
    return `Update deploy-azure-storage.json schema URL`;
  };

  get resolution(): string {
    return `{
  "$schema": "${this.schema}"
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/deploy-azure-storage.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.deployAzureStorageJson) {
      return;
    }

    if (project.deployAzureStorageJson.$schema !== this.schema) {
      this.addFinding(findings);
    }
  }
}