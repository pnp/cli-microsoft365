import * as path from 'path';
import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding, Occurrence } from '../../report-model';

export class FN014009_CODE_launch_hostedWorkbench_url extends JsonRule {
  constructor(private url: string) {
    super();
  }

  get id(): string {
    return 'FN014009';
  }

  get title(): string {
    return 'Hosted workbench URL in .vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode/launch.json file, update the url property for the hosted workbench launch configuration`;
  }

  get resolution(): string {
    return `{
  "configurations": [
    {
      "url": "${this.url}"
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
        configuration.url !== this.url) {
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