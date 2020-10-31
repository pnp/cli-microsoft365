import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012002_TSC_moduleResolution extends Rule {
  constructor(private moduleResolution: string) {
    super();
  }

  get id(): string {
    return 'FN012002';
  }

  get title(): string {
    return 'tsconfig.json moduleResolution';
  }

  get description(): string {
    return `Update moduleResolution in tsconfig.json`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "moduleResolution": "${this.moduleResolution}"
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './tsconfig.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson || !project.tsConfigJson.compilerOptions) {
      return;
    }

    if (project.tsConfigJson.compilerOptions.moduleResolution !== this.moduleResolution) {
      this.addFinding(findings);
    }
  }
}