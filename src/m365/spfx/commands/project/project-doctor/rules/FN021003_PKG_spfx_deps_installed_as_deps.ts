import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as spfxDeps from '../spfx-deps';

export class FN021003_PKG_spfx_deps_installed_as_deps extends JsonRule {
  get id(): string {
    return 'FN021003';
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
      !project.packageJson.devDependencies) {
      return;
    }

    const projectDevDeps = Object.keys(project.packageJson.devDependencies);
    projectDevDeps.forEach(dep => {
      if (!spfxDeps.deps.includes(dep)) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, `devDependencies.${dep}`);
      this.addFindingWithCustomInfo(
        `${dep} installed as devDependency`,
        `Package ${dep} is installed as a devDependency. Install it as a dependency instead`,
        [{
          file: this.file,
          resolution: `install ${dep}@${project.version}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}