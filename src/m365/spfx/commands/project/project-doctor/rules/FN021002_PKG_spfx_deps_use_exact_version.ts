import { valid } from 'semver';
import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as spfxDeps from '../spfx-deps';

export class FN021002_PKG_spfx_deps_use_exact_version extends JsonRule {
  get id(): string {
    return 'FN021002';
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
    if (!project.version || !project.packageJson) {
      return;
    }

    const allSpfxDeps = spfxDeps.deps.concat(spfxDeps.devDeps);

    if (project.packageJson.dependencies) {
      const projectDeps = Object.keys(project.packageJson.dependencies);
      this.validateDependencies({
        dependencies: projectDeps,
        isDevDep: false,
        allSpfxDeps,
        project,
        findings
      });
    }

    if (project.packageJson.devDependencies) {
      const projectDevDeps = Object.keys(project.packageJson.devDependencies);
      this.validateDependencies({
        dependencies: projectDevDeps,
        isDevDep: true,
        allSpfxDeps,
        project,
        findings
      });
    }
  }

  private validateDependencies({ dependencies, isDevDep, allSpfxDeps, project, findings }: { dependencies: string[], isDevDep: boolean, allSpfxDeps: string[], project: Project, findings: Finding[] }): void {
    dependencies.forEach(dep => {
      const depVersion = isDevDep ?
        project.packageJson!.devDependencies![dep] :
        project.packageJson!.dependencies![dep];

      if (!allSpfxDeps.includes(dep) ||
        valid(depVersion)) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, `${isDevDep ? 'devDependencies' : 'dependencies'}.${dep}`);
      this.addFindingWithCustomInfo(
        `${dep} is not using exact version`,
        `${dep} is referenced using a range ${depVersion}. Install the exact version matching the project ${dep}@${project.version}`,
        [{
          file: this.file,
          resolution: `${isDevDep ? 'installDev' : 'install'} ${dep}@${project.version}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}