import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";
import { FN018002_TEAMS_manifest } from "./FN018002_TEAMS_manifest";

export class FN018005_TEAMS_deprecated_manifest extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN018005';
  }

  get supersedes(): string[] {
    return ['FN018002'];
  }

  get title(): string {
    // title is kept empty so that the finding isn't reported
    return '';
  }

  get description(): string {
    return `Manually creating Microsoft Teams manifests for web parts is no longer necessary because they're created automatically`;
  }

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'cmd';
  }

  get file(): string {
    return '';
  };

  get severity(): string {
    return 'Optional';
  }

  visit(project: Project, findings: Finding[]): void {
    // this rule should be applied whenever the FN018002_TEAMS_manifest is
    const deprecatedRule: Rule = new FN018002_TEAMS_manifest();
    const fn018002Findings: Finding[] = [];
    deprecatedRule.visit(project, fn018002Findings);

    if (fn018002Findings.length > 0) {
      this.addFinding(findings);
    }
  }
}
