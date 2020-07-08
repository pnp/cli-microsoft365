import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN012005_TSC_typeRoots_microsoft extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012005';
  }

  get title(): string {
    return 'tsconfig.json typeRoots ./node_modules/@microsoft';
  }

  get description(): string {
    return `Add ./node_modules/@microsoft to typeRoots in tsconfig.json`;
  };

  get resolution(): string {
    return `{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
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

    if (!project.tsConfigJson.compilerOptions.typeRoots ||
      project.tsConfigJson.compilerOptions.typeRoots.indexOf('./node_modules/@microsoft') < 0) {
      this.addFinding(findings);
    }
  }
}