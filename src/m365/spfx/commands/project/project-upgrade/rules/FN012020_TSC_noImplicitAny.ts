import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012020_TSC_noImplicitAny extends JsonRule {
  constructor(private noImplicitAny: boolean) {
    super();
  }

  get id(): string {
    return 'FN012020';
  }

  get title(): string {
    return 'tsconfig.json noImplicitAny';
  }

  get description(): string {
    return `Add noImplicitAny in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "noImplicitAny": ${this.noImplicitAny}
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

    if (project.tsConfigJson.compilerOptions.noImplicitAny !== this.noImplicitAny) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.noImplicitAny');
      this.addFindingWithPosition(findings, node);
    }
  }
}
