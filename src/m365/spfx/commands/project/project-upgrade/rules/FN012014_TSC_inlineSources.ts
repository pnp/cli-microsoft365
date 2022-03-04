import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012014_TSC_inlineSources extends JsonRule {
  constructor(private inlineSources: boolean) {
    super();
  }

  get id(): string {
    return 'FN012014';
  }

  get title(): string {
    return 'tsconfig.json compiler options inlineSources';
  }

  get description(): string {
    return `Update tsconfig.json inlineSources value`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "inlineSources": ${this.inlineSources}
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

    if (project.tsConfigJson.compilerOptions.inlineSources !== this.inlineSources) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.inlineSources');
      this.addFindingWithPosition(findings, node);
    }
  }
}