import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN006005_CFG_PS_metadata extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN006005';
  }

  get title(): string {
    return 'package-solution.json metadata';
  }

  get description(): string {
    return `In package-solution.json add metadata section`;
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

    if (!project.packageSolutionJson.solution.metadata) {
      const solutionDescription = `${project.packageJson?.name} description`;
      const resolution = {
        solution: {
          metadata: {
            shortDescription: {
              default: solutionDescription
            },
            longDescription: {
              default: solutionDescription
            },
            screenshotPaths: [],
            videoUrl: '',
            categories: []
          }
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