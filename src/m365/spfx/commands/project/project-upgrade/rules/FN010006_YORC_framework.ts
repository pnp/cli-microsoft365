import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN010006_YORC_framework extends Rule {
  constructor(private framework: string, private add: boolean) {
    super();
  }

  get id(): string {
    return 'FN010006';
  }

  get title(): string {
    return '.yo-rc.json framework';
  }

  get description(): string {
    return `${this.add ? 'Update' : 'Remove'} framework in .yo-rc.json`;
  };

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "framework": "${this.framework}"
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

    if (this.add) {
      if (project.yoRcJson["@microsoft/generator-sharepoint"].framework !== this.framework) {
        this.addFinding(findings);
      }
    }
    else {
      if (project.yoRcJson["@microsoft/generator-sharepoint"].framework) {
        this.addFinding(findings);
      }
    }
  }
}