import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN023001_GITIGNORE_release extends Rule {
  constructor(private add: boolean = true) {
    super();
  }

  get id(): string {
    return 'FN023001';
  }

  get title(): string {
    return `.gitignore 'release' folder`;
  }

  get description(): string {
    if (this.add) {
      return `To .gitignore add the 'release' folder`;
    }
    else {
      return `From .gitignore remove the 'release' folder`;
    }
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

    if (/^release$/m.test(project.gitignore.source) !== this.add) {
      this.addFinding(findings);
    }
  }
}
