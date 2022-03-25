import * as path from 'path';
import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding, Occurrence } from '../../report-model';

export class FN014006_CODE_launch_sourceMapPathOverrides extends JsonRule {
  constructor(private overrideKey: string, private overrideValue: string) {
    super();
  }

  get id(): string {
    return 'FN014006';
  }

  get title(): string {
    return 'sourceMapPathOverrides in .vscode/launch.json';
  }

  get description(): string {
    return `In the .vscode/launch.json file, for each configuration, in the sourceMapPathOverrides property, add "${this.overrideKey}": "${this.overrideValue}"`;
  }

  get resolution(): string {
    return `{
  "configurations": [
    {
      "sourceMapPathOverrides": {
        "${this.overrideKey}": "${this.overrideValue}"
      }
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
      if (configuration.sourceMapPathOverrides &&
        !configuration.sourceMapPathOverrides[this.overrideKey]) {
        const node = this.getAstNodeFromFile(project.vsCode!.launchJson!, `configurations[${i}].sourceMapPathOverrides`);
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