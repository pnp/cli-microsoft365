import { Finding } from "../../report-model/index.js";
import { Project } from "../../project-model/index.js";
import { Rule } from '../../Rule.js';

export class FN023003_GITIGNORE_libdts extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023003';
  }

  get title(): string {
    return `.gitignore 'lib-dts' folder`;
  }

  get description(): string {
    return `To .gitignore add the 'lib-dts' folder`;
  }

  get resolution(): string {
    return `lib-dts`;
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

    if (!/^lib-dts$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
