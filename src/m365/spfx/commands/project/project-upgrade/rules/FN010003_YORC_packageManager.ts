import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN010003_YORC_packageManager extends JsonRule {
  constructor(private packageManager: string) {
    super();
  }

  get id(): string {
    return 'FN010003';
  }

  get title(): string {
    return '.yo-rc.json packageManager';
  }

  get description(): string {
    return `Update packageManager in .yo-rc.json`;
  }

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "packageManager": "${this.packageManager}"
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Recommended';
  }

  get file(): string {
    return './.yo-rc.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.yoRcJson) {
      return;
    }

    if (project.yoRcJson["@microsoft/generator-sharepoint"].packageManager !== this.packageManager) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.packageManager');
      this.addFindingWithPosition(findings, node);
    }
  }
}