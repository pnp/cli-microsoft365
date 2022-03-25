import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN021005_PKG_types_installed_as_devdep extends JsonRule {
  get id(): string {
    return 'FN021005';
  }

  get title(): string {
    return '';
  }

  get description(): string {
    return '';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  get resolutionType(): string {
    return 'cmd';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.version ||
      !project.packageJson ||
      !project.packageJson.dependencies) {
      return;
    }

    const projectDeps = Object.keys(project.packageJson.dependencies);
    projectDeps.forEach(dep => {
      if (dep.indexOf('@types/') < 0) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, `dependencies.${dep}`);
      this.addFindingWithCustomInfo(
        `${dep} installed as a dependency`,
        `Package ${dep} is installed as a dependency. Install it as a devDependency instead`,
        [{
          file: this.file,
          resolution: `installDev ${dep}@${project.packageJson!.dependencies![dep]}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}