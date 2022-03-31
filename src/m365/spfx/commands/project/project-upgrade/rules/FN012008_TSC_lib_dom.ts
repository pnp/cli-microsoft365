import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012008_TSC_lib_dom extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012008';
  }

  get title(): string {
    return 'tsconfig.json dom lib';
  }

  get description(): string {
    return `Add dom lib in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "lib": [
      "dom"
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
      project.tsConfigJson.compilerOptions.lib.indexOf('dom') < 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
      this.addFindingWithPosition(findings, node);
    }
  }
}