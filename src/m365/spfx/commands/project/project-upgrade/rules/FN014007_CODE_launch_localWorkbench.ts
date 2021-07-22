import * as path from 'path';
import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { JsonRule } from './JsonRule';

export class FN014007_CODE_launch_localWorkbench extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN014007';
  }

  get title(): string {
    return 'Local workbench in .vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode/launch.json file, remove the local workbench launch configuration`;
  }

  get resolution(): string {
    return ``;
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
      if (configuration.url &&
        configuration.url.indexOf('/temp/workbench.html') > -1) {
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