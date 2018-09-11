import { Finding } from "../";
import { Project } from "../model";
import { Rule } from "./Rule";

export class FN006003_CFG_PS_skipFeatureDeployment extends Rule {
  constructor(private valueType: string) {
    super();
  }

  get id(): string {
    return 'FN006003';
  }

  get title(): string {
    return 'package-solution.json skipFeatureDeployment';
  }

  get description(): string {
    return `In file package-solution.json update the type of the value of the skipFeatureDeployment property`;
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/package-solution.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    const valueType: string = typeof project.packageSolutionJson.solution.skipFeatureDeployment;

    if (valueType !== 'undefined' &&
      valueType !== this.valueType) {
      const resolution: string = `{
  "solution": {
    "skipFeatureDeployment": ${(this.valueType === 'string' ? `"${project.packageSolutionJson.solution.skipFeatureDeployment}"` : `${project.packageSolutionJson.solution.skipFeatureDeployment}`)}
  }
}`;

      this.addFindingWithCustomInfo(this.title, this.description, [{
        file: this.file,
        resolution: resolution
      }], findings);
    }
  }
}