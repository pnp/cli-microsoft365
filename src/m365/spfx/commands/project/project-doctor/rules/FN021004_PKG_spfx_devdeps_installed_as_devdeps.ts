import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as spfxDeps from '../spfx-deps';

export class FN021004_PKG_spfx_devdeps_installed_as_devdeps extends JsonRule {
  get id(): string {
    return 'FN021004';
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
      if (!spfxDeps.devDeps.includes(dep)) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, `dependencies.${dep}`);
      this.addFindingWithCustomInfo(
        `${dep} installed as a dependency`,
        `Package ${dep} is installed as a dependency. Install it as a devDependency instead`,
        [{
          file: this.file,
          resolution: `installDev ${dep}@${project.version}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}