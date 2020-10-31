import { Finding } from "../";
import { Project, VsCodeSettingsJsonJsonSchema } from "../../model";
import { Rule } from "./Rule";

export class FN014004_CODE_settings_jsonSchemas_configJson_url extends Rule {
  constructor(private url: string) {
    super();
  }

  get id(): string {
    return 'FN014004';
  }

  get title(): string {
    return 'URL of the config.json JSON schema in .vscode/settings.json';
  }

  get description(): string {
    return `Update the URL of the config.json JSON schema in .vscode/settings.json`;
  };

  get resolution(): string {
    return `{
  "json.schemas": [
    {
      "fileMatch": [
        "/config/config.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/configJson/schemas/config-v1.schema.json"
    }
  ]
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
    if (!project.vsCode ||
      !project.vsCode.settingsJson ||
      !project.vsCode.settingsJson["json.schemas"]) {
      return;
    }

    const schemas: VsCodeSettingsJsonJsonSchema[] = project.vsCode.settingsJson["json.schemas"] as VsCodeSettingsJsonJsonSchema[];
    for (let i: number = 0; i < schemas.length; i++) {
      const schema: VsCodeSettingsJsonJsonSchema = schemas[i];
      if (schema.fileMatch.indexOf('/config/config.json') === -1) {
        continue;
      }
      
      if (schema.url !== this.url) {
        this.addFinding(findings);
      }

      return;
    }
  }
}