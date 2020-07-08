import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN014001_CODE_settings_jsonSchemas extends Rule {
  constructor(private add: boolean) {
    super();
  }

  get id(): string {
    return 'FN014001';
  }

  get title(): string {
    return 'json.schemas in .vscode/settings.json';
  }

  get description(): string {
    return `${(this.add ? 'Add' : 'Remove')} json.schemas in .vscode/settings.json`;
  };

  get resolution(): string {
    return `{
  "json.schemas": []
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return '.vscode/settings.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode || !project.vsCode.settingsJson) {
      return;
    }

    if (this.add) {
      if (!project.vsCode.settingsJson["json.schemas"]) {
        this.addFinding(findings);
      }
    }
    else {
      if (project.vsCode.settingsJson["json.schemas"]) {
        this.addFinding(findings);
      }
    }
  }
}