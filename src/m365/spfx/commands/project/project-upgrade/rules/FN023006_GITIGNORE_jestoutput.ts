import { Finding } from "../../report-model/index.js";
import { Project } from "../../project-model/index.js";
import { Rule } from '../../Rule.js';

export class FN023006_GITIGNORE_jestoutput extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023006';
  }

  get title(): string {
    return `.gitignore 'jest-output' folder`;
  }

  get description(): string {
    return `To .gitignore add the 'jest-output' folder`;
  }

  get resolution(): string {
    return `jest-output`;
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

    if (!/^jest-output$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
