import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012003_TSC_skipLibCheck extends JsonRule {
  private skipLibCheck: boolean;

  constructor(options: { skipLibCheck: boolean }) {
    super();
    this.skipLibCheck = options.skipLibCheck;
  }

  get id(): string {
    return 'FN012003';
  }

  get title(): string {
    return 'tsconfig.json skipLibCheck';
  }

  get description(): string {
    return `Update skipLibCheck in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "skipLibCheck": true
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

    if (project.tsConfigJson.compilerOptions.skipLibCheck !== this.skipLibCheck) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.skipLibCheck');
      this.addFindingWithPosition(findings, node);
    }
  }
}