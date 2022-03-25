import { satisfies } from 'semver';
import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as spfxDeps from '../spfx-deps';

export class FN021001_PKG_spfx_deps_versions_match_project_version extends JsonRule {
  constructor() {
    super();  
  }

  get id(): string {
    return 'FN021001';
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
        satisfies(project.version as string, depVersion)) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, `${isDevDep ? 'devDependencies' : 'dependencies'}.${dep}`);
      this.addFindingWithCustomInfo(
        `${dep} doesn't match project version`,
        `${dep}@${depVersion} doesn't match the project version ${project.version}`,
        [{
          file: this.file,
          resolution: `${isDevDep ? 'installDev' : 'install'} ${dep}@${project.version}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}