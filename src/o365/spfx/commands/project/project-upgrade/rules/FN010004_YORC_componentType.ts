import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN010004_YORC_componentType extends Rule {
  constructor(private componentType: string) {
    super();
  }

  get id(): string {
    return 'FN010004';
  }

  get title(): string {
    return '.yo-rc.json componentType';
  }

  get description(): string {
    return `Update componentType in .yo-rc.json`;
  };

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "componentType": "${this.componentType}"
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"].componentType !== this.componentType) {
      this.addFinding(findings);
    }
  }
}