import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN010001_YORC_version extends JsonRule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN010001';
  }

  get title(): string {
    return '.yo-rc.json version';
  }

  get description(): string {
    return `Update version in .yo-rc.json`;
  };

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "version": "${this.version}"
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return './.yo-rc.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.yoRcJson) {
      return;
    }

    if (project.yoRcJson["@microsoft/generator-sharepoint"].version !== this.version) {
      const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.version');
      this.addFindingWithPosition(findings, node);
    }
  }
}