import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012001_TSC_module extends JsonRule {
  constructor(private module: string) {
    super();
  }

  get id(): string {
    return 'FN012001';
  }

  get title(): string {
    return 'tsconfig.json module';
  }

  get description(): string {
    return `Update module type in tsconfig.json`;
  }

  get resolution(): string {
    return `{
  "compilerOptions": {
    "module": "${this.module}"
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

    if (project.tsConfigJson.compilerOptions.module !== this.module) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.module');
      this.addFindingWithPosition(findings, node);
    }
  }
}