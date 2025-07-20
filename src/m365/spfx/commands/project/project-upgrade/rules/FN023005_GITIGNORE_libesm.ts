import { Finding } from "../../report-model/index.js";
import { Project } from "../../project-model/index.js";
import { Rule } from '../../Rule.js';

export class FN023005_GITIGNORE_libesm extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023005';
  }

  get title(): string {
    return `.gitignore 'lib-esm' folder`;
  }

  get description(): string {
    return `To .gitignore add the 'lib-esm' folder`;
  }

  get resolution(): string {
    return `lib-esm`;
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

    if (!/^lib-esm$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
