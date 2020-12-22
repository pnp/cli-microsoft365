import { Finding } from "../";
import { Project, TsConfigJson } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN012013_TSC_exclude extends JsonRule {
  constructor(private exclude: string[]) {
    super();
  }

  get id(): string {
    return 'FN012013';
  }

  get title(): string {
    return 'tsconfig.json exclude property';
  }

  get description(): string {
    return `Update tsconfig.json exclude property`;
  };

  get resolution(): string {
    return JSON.stringify({
      exclude: this.exclude
    }, null, 2);
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './tsconfig.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson) {
      return;
    }

    if (!project.tsConfigJson.exclude ||
      this.exclude.filter(e => ((project.tsConfigJson as TsConfigJson).exclude as string[]).indexOf(e) < 0).length > 0) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'exclude');
      this.addFindingWithPosition(findings, node);
    }
  }
}