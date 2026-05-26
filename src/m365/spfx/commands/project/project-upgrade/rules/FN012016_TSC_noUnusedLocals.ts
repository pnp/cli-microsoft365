import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012016_TSC_noUnusedLocals extends JsonRule {
  private noUnusedLocals: boolean;

  constructor(options: { noUnusedLocals: boolean }) {
    super();
    this.noUnusedLocals = options.noUnusedLocals;
  }

  get id(): string {
    return 'FN012016';
  }

  get title(): string {
    return 'tsconfig.json compiler options noUnusedLocals';
  }

  get description(): string {
    return `Update tsconfig.json noUnusedLocals value`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "noUnusedLocals": ${this.noUnusedLocals}
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

    if (project.tsConfigJson.compilerOptions.noUnusedLocals !== this.noUnusedLocals) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.noUnusedLocals');
      this.addFindingWithPosition(findings, node);
    }
  }
}