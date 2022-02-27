import * as path from 'path';
import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { JsonRule } from './JsonRule';

export class FN014008_CODE_launch_hostedWorkbench_type extends JsonRule {
  constructor(private type: string) {
    super();
  }

  get id(): string {
    return 'FN014008';
  }

  get title(): string {
    return 'Hosted workbench type in .vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration`;
  }

  get resolution(): string {
    return `{
  "configurations": [
    {
      "type": "${this.type}"
    }
  ]
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Recommended';
  }

  get file(): string {
    return '.vscode/launch.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.vsCode ||
      !project.vsCode.launchJson ||
      !project.vsCode.launchJson.configurations) {
      return;
    }

    const occurrences: Occurrence[] = [];
    project.vsCode.launchJson.configurations.forEach((configuration, i) => {
      if (configuration.name === 'Hosted workbench' &&
        configuration.type !== this.type) {
        const node = this.getAstNodeFromFile(project.vsCode!.launchJson!, `configurations[${i}].url`);
        occurrences.push({
          file: path.relative(project.path, this.file),
          resolution: this.resolution,
          position: this.getPositionFromNode(node)
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}