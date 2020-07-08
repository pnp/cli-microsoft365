import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN010002_YORC_isCreatingSolution extends Rule {
  constructor(private value: boolean) {
    super();
  }

  get id(): string {
    return 'FN010002';
  }

  get title(): string {
    return '.yo-rc.json isCreatingSolution';
  }

  get description(): string {
    return `Update isCreatingSolution in .yo-rc.json`;
  };

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "isCreatingSolution": ${this.value.toString()}
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"].isCreatingSolution !== this.value) {
      this.addFinding(findings);
    }
  }
}