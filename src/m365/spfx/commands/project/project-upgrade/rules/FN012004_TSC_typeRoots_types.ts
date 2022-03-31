import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012004_TSC_typeRoots_types extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012004';
  }

  get title(): string {
    return 'tsconfig.json typeRoots ./node_modules/@types';
  }

  get description(): string {
    return `Add ./node_modules/@types to typeRoots in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@types"
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

    if (!project.tsConfigJson.compilerOptions.typeRoots ||
      project.tsConfigJson.compilerOptions.typeRoots.indexOf('./node_modules/@types') < 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.typeRoots');
      this.addFindingWithPosition(findings, node);
    }
  }
}