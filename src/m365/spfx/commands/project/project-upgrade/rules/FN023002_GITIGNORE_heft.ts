import { Finding } from "../../report-model/index.js";
import { Project } from "../../project-model/index.js";
import { Rule } from '../../Rule.js';

export class FN023002_GITIGNORE_heft extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023002';
  }

  get title(): string {
    return `.gitignore '.heft' folder`;
  }

  get description(): string {
    return `To .gitignore add the '.heft' folder`;
  }

  get resolution(): string {
    return `.heft`;
  }

  get resolutionType(): string {
    return 'text';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './.gitignore';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.gitignore) {
      return;
    }

    if (!/^\.heft$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
