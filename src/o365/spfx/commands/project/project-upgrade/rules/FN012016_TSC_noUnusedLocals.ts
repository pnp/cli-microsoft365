import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012016_TSC_noUnusedLocals extends Rule {
  constructor(private noUnusedLocals: boolean) {
    super();
  }

  get id(): string {
    return 'FN012016';
  }

  get title(): string {
    return 'tsconfig.json compiler options noUnusedLocals';
  }

  get description(): string {
    return `Update tsconfig.json noUnusedLocals value`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "noUnusedLocals": ${this.noUnusedLocals}
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

    if (project.tsConfigJson.compilerOptions.noUnusedLocals !== this.noUnusedLocals) {
      this.addFinding(findings);
    }
  }
}