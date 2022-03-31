import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012003_TSC_skipLibCheck extends JsonRule {
  constructor(private skipLibCheck: boolean) {
    super();
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