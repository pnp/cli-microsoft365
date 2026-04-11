import { Rule } from "../../Rule.js";
import { Project } from "../../project-model/index.js";
import { Finding } from "../../report-model/index.js";

export class FN021009_PKG_overrides_rushstack_heft extends Rule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN021009';
  }

  get title(): string {
    return 'package.json overrides.@rushstack/heft';
  }

  get description(): string {
    return 'Update package.json overrides.@rushstack/heft property';
  }

  get resolution(): string {
    return `npm pkg set overrides.@rushstack/heft="${this.version}"`;
  }

  get resolutionType(): string {
    return 'cmd';
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

    if (!project.packageJson.overrides ||
      typeof project.packageJson.overrides !== 'object' ||
      !project.packageJson.overrides['@rushstack/heft'] ||
      project.packageJson.overrides['@rushstack/heft'] !== this.version) {
      this.addFinding(findings);
    }
  }
}