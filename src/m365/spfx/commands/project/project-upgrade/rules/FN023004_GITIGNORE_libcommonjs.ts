import { Finding } from "../../report-model/index.js";
import { Project } from "../../project-model/index.js";
import { Rule } from '../../Rule.js';

export class FN023004_GITIGNORE_libcommonjs extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023004';
  }

  get title(): string {
    return `.gitignore 'lib-commonjs' folder`;
  }

  get description(): string {
    return `To .gitignore add the 'lib-commonjs' folder`;
  }

  get resolution(): string {
    return `lib-commonjs`;
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

    if (!/^lib-commonjs$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
