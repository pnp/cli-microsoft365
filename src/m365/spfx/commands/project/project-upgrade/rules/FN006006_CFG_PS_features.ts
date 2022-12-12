import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding, Occurrence } from '../../report-model';

export class FN006006_CFG_PS_features extends JsonRule {
  get id(): string {
    return 'FN006006';
  }

  get title(): string {
    return 'package-solution.json features';
  }

  get description(): string {
    return `In package-solution.json add features for components`;
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
      !project.packageSolutionJson.solution ||
      // if project already has features defined, we don't need to do anything
      (project.packageSolutionJson.solution.features && project.packageSolutionJson.solution.features.length > 0) ||
      // if there are no components, we don't need to do anything
      !project.manifests || project.manifests.length < 1) {
      return;
    }

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      const resolution = {
        solution: {
          features: [
            {
              title: `${project.packageJson?.name} ${manifest.alias} Feature`,
              description: `The feature that activates ${manifest.alias} from the ${project.packageJson?.name} solution.`,
              id: manifest.id,
              version: project.packageSolutionJson!.solution!.version,
              componentIds: [manifest.id]
            }
          ]
        }
      };
      const node = this.getAstNodeFromFile(project.packageSolutionJson!, 'solution');
      occurrences.push({
        file: this.file,
        resolution: JSON.stringify(resolution, null, 2),
        position: this.getPositionFromNode(node)
      });
    });

    this.addFindingWithOccurrences(occurrences, findings);
  }
}