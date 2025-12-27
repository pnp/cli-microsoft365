import { JsonRule } from '../../JsonRule.js';
import { Project, TsConfigJson } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012013_TSC_exclude extends JsonRule {
  private exclude: string[];
  private add: boolean;

  constructor(options: { exclude: string[]; add?: boolean }) {
    super();
    this.exclude = options.exclude;
    this.add = options.add ?? true;
  }

  get id(): string {
    return 'FN012013';
  }

  get title(): string {
    return 'tsconfig.json exclude property';
  }

  get description(): string {
    if (this.add) {
      return `Update tsconfig.json exclude property`;
    }
    else {
      return `Remove tsconfig.json exclude property`;
    }
  }

  get resolution(): string {
    return JSON.stringify({
      exclude: this.exclude
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

    if (this.add) {
      if (!project.tsConfigJson.exclude ||
        this.exclude.filter(e => ((project.tsConfigJson as TsConfigJson).exclude as string[]).indexOf(e) < 0).length > 0) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'exclude');
        this.addFindingWithPosition(findings, node);
      }
    }
    else {
      if (project.tsConfigJson.exclude) {
        const node = this.getAstNodeFromFile(project.tsConfigJson, 'exclude');
        this.addFindingWithPosition(findings, node);
      }
    }
  }
}