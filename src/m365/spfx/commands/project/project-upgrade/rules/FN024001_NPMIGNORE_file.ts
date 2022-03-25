import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { Rule } from '../../Rule';

export class FN024001_NPMIGNORE_file extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN024001';
  }

  get title(): string {
    return `Create .npmignore`;
  }

  get description(): string {
    return `Create the .npmignore file`;
  }

  get resolution(): string {
    return `!dist
config

gulpfile.js

release
src
temp

tsconfig.json
tslint.json

*.log

.yo-rc.json
.vscode
`;
  }

  get resolutionType(): string {
    return 'text';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './.npmignore';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.npmignore) {
      this.addFinding(findings);
    }
  }
}
