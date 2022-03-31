import { JsonRule } from '../../JsonRule';
import { PackageJson, Project } from '../../project-model';
import { Finding } from '../../report-model';
import * as spfxDeps from '../spfx-deps';

export class FN021008_PKG_no_duplicate_deps extends JsonRule {
  get id(): string {
    return 'FN021008';
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
      !project.packageJson.dependencies ||
      !project.packageJson.devDependencies) {
      return;
    }

    const projectDeps = Object.keys(project.packageJson.dependencies);
    const projectDevDeps = Object.keys(project.packageJson.devDependencies);

    const duplicateDeps = projectDeps.filter(dep => projectDevDeps.includes(dep));
    if (duplicateDeps.length === 0) {
      return;
    }

    duplicateDeps.forEach(dep => {
      const isDevDep = this.isDevDep(dep);
      const nodeToUninstall = this.getAstNodeFromFile(project.packageJson as PackageJson, `${isDevDep ? 'dependencies' : 'devDependencies'}.${dep}`);
      this.addFindingWithCustomInfo(
        `Duplicate ${dep} installed in the project`,
        `Duplicate ${dep} installed in the project. Install it only as a ${isDevDep ? 'devDependency' : 'dependency'}`,
        [{
          file: this.file,
          resolution: `${isDevDep ? 'installDev' : 'install'} ${dep}@${isDevDep ? project.packageJson!.devDependencies![dep] : project.packageJson!.dependencies![dep]}`,
          position: this.getPositionFromNode(nodeToUninstall)
        }], findings);
    });
  }

  private isDevDep(dep: string): boolean {
    if (dep.indexOf('@types/') === 0) {
      return true;
    }

    if (spfxDeps.devDeps.includes(dep)) {
      return true;
    }

    return false;
  }
}