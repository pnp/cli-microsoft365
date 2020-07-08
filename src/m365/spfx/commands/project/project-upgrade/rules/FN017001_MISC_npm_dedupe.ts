import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN017001_MISC_npm_dedupe extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN017001';
  }

  get title(): string {
    return 'Run npm dedupe';
  }

  get description(): string {
    return `If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.`;
  }

  get resolution(): string {
    return 'npm dedupe';
  };

  get resolutionType(): string {
    return 'cmd';
  }

  get file(): string {
    return './package.json';
  };

  get severity(): string {
    return 'Optional';
  }

  visit(project: Project, findings: Finding[]): void {
    const npmFinding: Finding | undefined = findings.find(f => typeof f.occurrences.find(o => o.resolution.indexOf('install') === 0 || o.resolution.indexOf('uninstall') === 0) !== 'undefined');
    if (npmFinding) {
      this.addFinding(findings);
    }
  }
}
