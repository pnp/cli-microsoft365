import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012015_TSC_strictNullChecks extends Rule {
  constructor(private strictNullChecks: boolean) {
    super();
  }

  get id(): string {
    return 'FN012015';
  }

  get title(): string {
    return 'tsconfig.json compiler options strictNullChecks';
  }

  get description(): string {
    return `Update tsconfig.json strictNullChecks value`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "strictNullChecks": ${this.strictNullChecks}
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

    if (project.tsConfigJson.compilerOptions.strictNullChecks !== this.strictNullChecks) {
      this.addFinding(findings);
    }
  }
}