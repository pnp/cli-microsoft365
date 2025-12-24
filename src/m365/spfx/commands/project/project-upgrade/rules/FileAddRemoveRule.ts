import fs from 'fs';
import path from 'path';
import { Project } from '../../project-model/index.js';
import { Finding } from "../../report-model/Finding.js";
import { Rule } from '../../Rule.js';

export interface FileAddRemoveRuleOptions {
  filePath: string;
  add: boolean;
  contents?: string;
}

export abstract class FileAddRemoveRule extends Rule {
  protected filePath: string;
  protected add: boolean;
  protected contents?: string;

  constructor(options: FileAddRemoveRuleOptions) {
    super();
    this.filePath = path.normalize(options.filePath);
    this.add = options.add;
    this.contents = options.contents;
  }

  get title(): string {
    return this.filePath;
  }

  get description(): string {
    return `${this.add ? 'Add' : 'Remove'} file ${this.filePath}`;
  }

  get resolution(): string {
    if (this.add) {
      return `add_cmd[BEFOREPATH]"${this.filePath}"[AFTERPATH][BEFORECONTENT]
${this.contents}
[AFTERCONTENT]`;
    }
    else {
      return `remove_cmd "${this.filePath}"`;
    }
  }

  get resolutionType(): string {
    return 'cmd';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return this.filePath;
  }

  public visit(project: Project, notifications: Finding[]): void {
    const targetPath: string = path.join(project.path, this.filePath);
    if ((!this.add && fs.existsSync(targetPath)) ||
      (this.add && !fs.existsSync(targetPath))) {
      this.addFinding(notifications);
      return;
    }
    if (this.add && this.contents) {
      const fileContent: string = fs.readFileSync(path.join(project.path, this.filePath), 'utf8');
      if (fileContent !== this.contents) {
        this.addFinding(notifications);
      }
    }
  }
}
