import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN014003_CODE_launch extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN014003';
  }

  get title(): string {
    return '.vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode folder, add the launch.json file`;
  };

  get resolution(): string {
    return `{
  /**
    Install Chrome Debugger Extension for Visual Studio Code
    to debug your components with the Chrome browser:
    https://aka.ms/spfx-debugger-extensions
    */
  "version": "0.2.0",
  "configurations": [{
      "name": "Local workbench",
      "type": "chrome",
      "request": "launch",
      "url": "https://localhost:4321/temp/workbench.html",
      "webRoot": "\${workspaceRoot}",
      "sourceMaps": true,
      "sourceMapPathOverrides": {
        "webpack:///../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../../src/*": "\${webRoot}/src/*"
      },
      "runtimeArgs": [
        "--remote-debugging-port=9222"
      ]
    },
    {
      "name": "Hosted workbench",
      "type": "chrome",
      "request": "launch",
      "url": "https://enter-your-SharePoint-site/_layouts/workbench.aspx",
      "webRoot": "\${workspaceRoot}",
      "sourceMaps": true,
      "sourceMapPathOverrides": {
        "webpack:///../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../../src/*": "\${webRoot}/src/*"
      },
      "runtimeArgs": [
        "--remote-debugging-port=9222",
        "-incognito"
      ]
    }
  ]
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recommended';
  };

  get file(): string {
    return '.vscode/launch.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode || !project.vsCode.launchJson) {
      this.addFinding(findings);
    }
  }
}