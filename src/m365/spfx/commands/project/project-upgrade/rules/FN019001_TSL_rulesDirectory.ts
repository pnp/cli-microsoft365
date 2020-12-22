import { Project } from "../../model";
import { Finding } from "../Finding";
import { JsonRule } from "./JsonRule";

export class FN019001_TSL_rulesDirectory extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN019001';
  }

  get title(): string {
    return 'tslint.json rulesDirectory';
  }

  get description(): string {
    return `Remove rulesDirectory from tslint.json`;
  };

  get resolution(): string {
    return `{
  "rulesDirectory": []
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './tslint.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsLintJsonRoot) {
      return;
    }

    if (project.tsLintJsonRoot.rulesDirectory) {
      const node = this.getAstNodeFromFile(project.tsLintJsonRoot, 'rulesDirectory');
      this.addFindingWithPosition(findings, node);
    }
  }
}