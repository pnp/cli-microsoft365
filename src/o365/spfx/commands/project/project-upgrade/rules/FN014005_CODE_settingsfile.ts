import * as fs from "fs";
import { Finding } from "../";
import { Project } from "../../model";
import * as path from "path";
import { Rule } from "./Rule";

export class FN014005_CODE_settingsfile extends Rule {
  get id(): string {
    return 'FN014005';
  }

  get title(): string {
    return 'Missing vscode settings file';
  }

  get description(): string {
    return `Create file ${this.file} with provided content`;
  }

  get resolution(): string {
    return `// Place your settings in this file to overwrite default and user settings.
{
  // Configure glob patterns for excluding files and folders in the file explorer.
  "files.exclude": {
    "**/.git": true,
    "**/.DS_Store": true,
    "**/bower_components": true,
    "**/coverage": true,
    "**/lib-amd": true,
    "src/**/*.scss.ts": true
  },
  "typescript.tsdk": ".\\node_modules\\typescript\\lib",
  "json.schemas": [
    {
      "fileMatch": [
        "/config/config.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-web/lib/schemas/config.schema.json"
    },
    {
      "fileMatch": [
        "/config/copy-assets.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/copyAssets/copy-assets.schema.json"
    },
    {
      "fileMatch": [
        "/config/deploy-azure-storage.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/deployAzureStorage/deploy-azure-storage.schema.json"
    },
    {
      "fileMatch": [
        "/config/package-solution.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/packageSolution/package-solution.schema.json"
    },
    {
      "fileMatch": [
        "/config/serve.json"
      ],
      "url": "./node_modules/@microsoft/gulp-core-build-serve/lib/serve.schema.json"
    },
    {
      "fileMatch": [
        "/config/tslint.json"
      ],
      "url": "./node_modules/@microsoft/gulp-core-build-typescript/lib/schemas/tslint.schema.json"
    },
    {
      "fileMatch": [
        "/config/write-manifests.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/writeManifests/write-manifests.schema.json"
    },
    {
      "fileMatch": [
        "/config/configure-webpack.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/configureWebpack/configure-webpack.schema.json"
    },
    {
      "fileMatch": [
        "/config/configure-external-bundling-webpack.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/configureWebpack/configure-webpack-external-bundling.schema.json"
    },
    {
      "fileMatch": [
        "/copy-static-assets.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/copyStaticAssets/copy-static-assets.schema.json"
    }
  ]
}`;
  };
  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './.vscode/settings.json';
  }

  visit(project: Project, findings: Finding[]): void {
    const targetPath: string = path.join(project.path, this.file);
    if (!fs.existsSync(targetPath)) {
      this.addFinding(findings);
    }
  }
}