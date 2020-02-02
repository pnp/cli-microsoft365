import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN010005_YORC_environment extends Rule {
  constructor(private environment: string) {
    super();
  }

  get id(): string {
    return 'FN010005';
  }

  get title(): string {
    return '.yo-rc.json environment';
  }

  get description(): string {
    return `Update environment in .yo-rc.json`;
  };

  get resolution(): string {
    return `{
  "@microsoft/generator-sharepoint": {
    "environment": "${this.environment}"
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

    if (project.yoRcJson["@microsoft/generator-sharepoint"].environment !== this.environment) {
      this.addFinding(findings);
    }
  }
}