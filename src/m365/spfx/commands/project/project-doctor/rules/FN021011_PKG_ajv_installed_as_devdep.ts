import { JsonRule } from '../../JsonRule';
import { PackageJson, Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN021011_PKG_ajv_installed_as_devdep extends JsonRule {
  get id(): string {
    return 'FN021011';
  }

  get title(): string {
    return 'ajv installed as a devDependency';
  }

  get description(): string {
    return 'ajv is installed as a dependency. Install it as a devDependency instead';
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
    if (!projectDeps.includes('ajv')) {
      return;
    }

    const node = this.getAstNodeFromFile(project.packageJson as PackageJson, `dependencies.ajv`);
    this.addFindingWithCustomInfo(
      this.title,
      this.description,
      [{
        file: this.file,
        resolution: `installDev ajv@${project.packageJson.dependencies['ajv']}`,
        position: this.getPositionFromNode(node)
      }],
      findings
    );
  }
}