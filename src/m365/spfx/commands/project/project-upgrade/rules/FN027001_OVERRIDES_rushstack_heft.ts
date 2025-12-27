import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { DependencyRule } from "./DependencyRule.js";

export class FN027001_OVERRIDES_rushstack_heft extends DependencyRule {
  constructor(options: { version: string }) {
    super({ packageName: '@rushstack/heft', packageVersion: options.version, isOverride: true });
  }

  get id(): string {
    return 'FN027001';
  }

  visit(project: Project, findings: Finding[]): void {
    // If an override entry for the package already exists in package.json,
    // emit an extra finding to remove the existing override first. This avoids
    // having to use a separate remove-override rule (e.g. FN027002) in the upgrade scripts.
    if (project.packageJson?.overrides?.[this.packageName] &&
      project.packageJson.overrides[this.packageName] !== this.packageVersion) {
      const node = this.getAstNodeFromFile(project.packageJson, `overrides.${this.packageName}`);
      findings.push({
        id: `${this.id}_REMOVE`,
        title: this.packageName,
        description: `Remove existing SharePoint Framework override dependency package ${this.packageName}`,
        occurrences: [{
          file: this.file,
          resolution: `removeOverride overrides.${this.packageName}`,
          position: this.getPositionFromNode(node)
        }],
        resolutionType: 'cmd',
        severity: 'Required',
        supersedes: []
      });
    }

    super.visit(project, findings);
  }
}