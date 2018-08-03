import * as path from "path";
import * as fs from "fs";
import { Rule } from "./Rule";
import { Project } from "../model";
import { Finding } from "../Finding";

export abstract class FileRemoveRule extends Rule {
  public constructor(protected filePath: string) {
    super();
  }
  get title(): string {
    return this.filePath;
  }
  get description(): string {
    return `Delete file ${this.filePath}`;
  }
  get resolution(): string {
    return `rm ${this.filePath}`;
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
    if (!this.filePath) {
      return;
    }
    const targetPath: string = path.join(project.path, this.filePath);
    if (fs.existsSync(targetPath)) {
      this.addFinding(notifications);
    }
  }
}
