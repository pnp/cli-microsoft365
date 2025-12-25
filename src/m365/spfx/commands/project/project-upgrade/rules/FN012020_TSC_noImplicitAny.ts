import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012020_TSC_noImplicitAny extends JsonRule {
  private noImplicitAny: boolean;
  constructor(options: { noImplicitAny: boolean }) {
    super();
    this.noImplicitAny = options.noImplicitAny;
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
