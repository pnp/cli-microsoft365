import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012007_TSC_lib_es5 extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012007';
  }

  get title(): string {
    return 'tsconfig.json es5 lib';
  }

  get description(): string {
    return `Add es5 lib in tsconfig.json`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "lib": [
      "es5"
    ]
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

    if (!project.tsConfigJson.compilerOptions.lib ||
      project.tsConfigJson.compilerOptions.lib.indexOf('es5') < 0) {
      this.addFinding(findings);
    }
  }
}