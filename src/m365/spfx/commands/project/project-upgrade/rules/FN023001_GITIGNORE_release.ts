import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { Rule } from '../../Rule';

export class FN023001_GITIGNORE_release extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN023001';
  }

  get title(): string {
    return `.gitignore 'release' folder`;
  }

  get description(): string {
    return `To .gitignore add the 'release' folder`;
  }

  get resolution(): string {
    return `release`;
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

    if (!/^release$/m.test(project.gitignore.source)) {
      this.addFinding(findings);
    }
  }
}
