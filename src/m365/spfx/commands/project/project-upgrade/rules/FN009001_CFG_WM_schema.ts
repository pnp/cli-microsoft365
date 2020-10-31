import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN009001_CFG_WM_schema extends Rule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN009001';
  }

  get title(): string {
    return 'write-manifests.json schema';
  }

  get description(): string {
    return `Update write-manifests.json schema URL`;
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
    return './config/write-manifests.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.writeManifestsJson) {
      return;
    }

    if (project.writeManifestsJson.$schema !== this.schema) {
      this.addFinding(findings);
    }
  }
}