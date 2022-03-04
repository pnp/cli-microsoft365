import { JsonRule } from '../../JsonRule';
import { PackageJson, Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN021010_PKG_gulp_installed_as_devdep extends JsonRule {
  get id(): string {
    return 'FN021010';
  }

  get title(): string {
    return 'gulp installed as a devDependency';
  }

  get description(): string {
    return 'gulp is installed as a dependency. Install it as a devDependency instead';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  get resolution(): string {
    return '';
  }

  get resolutionType(): string {
    return 'cmd';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson ||
      !project.packageJson.dependencies) {
      return;
    }

    const projectDeps = Object.keys(project.packageJson.dependencies);
    if (!projectDeps.includes('gulp')) {
      return;
    }

    const node = this.getAstNodeFromFile(project.packageJson as PackageJson, `dependencies.gulp`);
    this.addFindingWithCustomInfo(
      this.title,
      this.description,
      [{
        file: this.file,
        resolution: `installDev gulp@${project.packageJson.dependencies['gulp']}`,
        position: this.getPositionFromNode(node)
      }],
      findings
    );
  }
}