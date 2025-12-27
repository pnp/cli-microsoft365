import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012011_TSC_outDir extends JsonRule {
  private outDir: string;

  constructor(options: { outDir: string }) {
    super();
    this.outDir = options.outDir;
  }

  get id(): string {
    return 'FN012011';
  }

  get title(): string {
    return 'tsconfig.json compiler options outDir';
  }

  get description(): string {
    return `Update tsconfig.json outDir value`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "outDir": "${this.outDir}"
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

    if (project.tsConfigJson.compilerOptions.outDir !== this.outDir) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.outDir');
      this.addFindingWithPosition(findings, node);
    }
  }
}