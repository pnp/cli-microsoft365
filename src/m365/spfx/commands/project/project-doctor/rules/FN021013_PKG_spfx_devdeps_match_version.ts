import { satisfies } from 'semver';
import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import * as spfxDeps from '../spfx-deps.js';

export class FN021013_PKG_spfx_devdeps_match_version extends JsonRule {
  constructor(private version: string) {
    super();
  }

  get id(): string {
    return 'FN021013';
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
    if (!project.packageJson || !project.packageJson.devDependencies) {
      return;
    }

    if (project.packageJson.devDependencies) {
      const projectDevDeps = Object.keys(project.packageJson.devDependencies);
      this.validateDependencies({
        dependencies: projectDevDeps,
        spfxDeps: spfxDeps.devDeps,
        project,
        version: this.version,
        findings
      });
    }
  }

  private validateDependencies({ dependencies, spfxDeps, project, version, findings }: { dependencies: string[], spfxDeps: string[], project: Project, version: string, findings: Finding[] }): void {
    dependencies.forEach(dep => {
      const depVersion = project.packageJson!.devDependencies![dep];

      if (!spfxDeps.includes(dep) ||
        satisfies(version, depVersion)) {
        return;
      }

      const node = this.getAstNodeFromFile(project.packageJson!, 'devDependencies');
      this.addFindingWithCustomInfo(
        `${dep} doesn't match project version`,
        `${dep}@${depVersion} doesn't match the project version ${project.version}`,
        [{
          file: this.file,
          resolution: `installDev ${dep}@${project.version}`,
          position: this.getPositionFromNode(node)
        }], findings);
    });
  }
}