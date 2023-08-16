import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012010_TSC_experimentalDecorators extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN012010';
  }

  get title(): string {
    return 'tsconfig.json compiler options experimental decorators';
  }

  get description(): string {
    return `Enable tsconfig.json experimental decorators`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "experimentalDecorators": true
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

    if (!project.tsConfigJson.compilerOptions.experimentalDecorators) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.experimentalDecorators');
      this.addFindingWithPosition(findings, node);
    }
  }
}