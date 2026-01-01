import { JsonRule } from "../../JsonRule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021005_PKG_scripts_test extends JsonRule {
  private script: string;
  private add: boolean;

  constructor(options: { script: string; add?: boolean }) {
    super();
    this.script = options.script;
    this.add = options.add ?? true;
  }

  get id(): string {
    return 'FN021005';
  }

  get title(): string {
    return 'package.json scripts.test';
  }

  get description(): string {
    return `${this.add ? 'Update' : 'Remove'} package.json scripts.test property`;
  }

  get resolution(): string {
    return `{
  "scripts": {
    "test": ${this.add ? `"${this.script}"` : '""'}
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
    return './package.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }

    if (this.add) {
      if (!project.packageJson.scripts ||
        typeof project.packageJson.scripts !== 'object' ||
        !project.packageJson.scripts.test ||
        project.packageJson.scripts.test !== this.script) {
        const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test');
        this.addFindingWithPosition(findings, node);
      }
    }
    else {
      if (project.packageJson.scripts?.test === this.script) {
        const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test');
        this.addFindingWithPosition(findings, node);
      }
    }
  }
}