import { JsonRule } from '../../JsonRule';
import { Project, TsConfigJson } from '../../project-model';
import { Finding } from '../../report-model';

export class FN012012_TSC_include extends JsonRule {
  constructor(private include: string[]) {
    super();
  }

  get id(): string {
    return 'FN012012';
  }

  get title(): string {
    return 'tsconfig.json include property';
  }

  get description(): string {
    return `Add to the tsconfig.json include property`;
  }

  get resolution(): string {
    return JSON.stringify({
      include: this.include
    }, null, 2);
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
    if (!project.tsConfigJson) {
      return;
    }

    if (!project.tsConfigJson.include ||
      this.include.filter(i => ((project.tsConfigJson as TsConfigJson).include as string[]).indexOf(i) < 0).length > 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'include');
      this.addFindingWithPosition(findings, node);
    }
  }
}