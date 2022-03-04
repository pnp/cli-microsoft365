import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012006_TSC_types_es6_collections extends JsonRule {
  constructor(private add: boolean) {
    super();
  }

  get id(): string {
    return 'FN012006';
  }

  get title(): string {
    return 'tsconfig.json es6-collections types';
  }

  get description(): string {
    return `${(this.add ? 'Add' : 'Remove')} es6-collections type in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "types": [
      "es6-collections"
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

    if (this.add) {
      if (!project.tsConfigJson.compilerOptions.types ||
        project.tsConfigJson.compilerOptions.types.indexOf('es6-collections') < 0) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types');
        this.addFindingWithPosition(findings, node);
      }
    }
    else {
      if (project.tsConfigJson.compilerOptions.types &&
        project.tsConfigJson.compilerOptions.types.indexOf('es6-collections') > -1) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types[es6-collections]');
        this.addFindingWithPosition(findings, node);
      }
    }
  }
}