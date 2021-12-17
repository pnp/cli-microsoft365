import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";
import { v4 } from 'uuid';

export class FN006006_CFG_PS_features extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN006006';
  }

  get title(): string {
    return 'package-solution.json features';
  }

  get description(): string {
    return `In package-solution.json add features section`;
  }

  get resolution(): string {
    return '';
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './config/package-solution.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    if (!project.packageSolutionJson.solution.features) {
      const resolution = {
        solution: {
          features: [
            {
              title: `${project.packageJson?.name} Feature`,
              description: `The feature that activates elements of the ${project.packageJson?.name} solution.`,
              id: v4(),
              version: project.packageSolutionJson.solution.version
            }
          ]
        }
      };
      const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution');
      this.addFindingWithOccurrences([{
        file: this.file,
        resolution: JSON.stringify(resolution, null, 2),
        position: this.getPositionFromNode(node)
      }], findings);
    }
  }
}