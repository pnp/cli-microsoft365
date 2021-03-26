import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN012018_TSC_lib_es2015_promise extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012018';
  }

  get title(): string {
    return 'tsconfig.json es2015.promise lib';
  }

  get description(): string {
    return `Add es2015.promise lib in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "lib": [
      "es2015.promise"
    ]
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
    return './tsconfig.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson || !project.tsConfigJson.compilerOptions) {
      return;
    }

    if (!project.tsConfigJson.compilerOptions.lib ||
      project.tsConfigJson.compilerOptions.lib.indexOf('es2015.promise') < 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
      this.addFindingWithPosition(findings, node);
    }
  }
}