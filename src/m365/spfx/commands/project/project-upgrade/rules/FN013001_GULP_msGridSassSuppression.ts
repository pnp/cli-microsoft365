import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN013001_GULP_msGridSassSuppression extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN013001';
  }

  get title(): string {
    return 'gulpfile.js ms-Grid sass suppression';
  }

  get description(): string {
    return `Add suppression for ms-Grid sass warning`;
  };

  get resolution(): string {
    return "build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);";
  };

  get resolutionType(): string {
    return 'js';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return './gulpfile.js';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.gulpfileJs) {
      return;
    }

    if (project.gulpfileJs.src.indexOf(this.resolution) < 0) {
      this.addFinding(findings);
    }
  }
}