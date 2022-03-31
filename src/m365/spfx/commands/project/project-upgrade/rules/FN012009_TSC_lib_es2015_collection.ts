import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012009_TSC_lib_es2015_collection extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012009';
  }

  get title(): string {
    return 'tsconfig.json es2015.collection lib';
  }

  get description(): string {
    return `Add es2015.collection lib in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "lib": [
      "es2015.collection"
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
      project.tsConfigJson.compilerOptions.lib.indexOf('es2015.collection') < 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
      this.addFindingWithPosition(findings, node);
    }
  }
}