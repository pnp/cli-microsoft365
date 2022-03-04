import { JsonRule } from '../../JsonRule';
import { PackageJson, Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN021012_PKG_no_duplicate_pnpjs_deps extends JsonRule {
  get id(): string {
    return 'FN021012';
  }

  get title(): string {
    return 'sp-pnp-js and @pnp/js installed in the project';
  }

  get description(): string {
    return 'sp-pnp-js and @pnp/js installed in the project. Consider uninstalling the deprecated sp-pnp-js package';
  }

  get severity(): string {
    return 'Optional';
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
    if (!project.packageJson) {
      return;
    }

    const projectDeps: string[] = [];
    const projectDevDeps: string[] = [];
    if (project.packageJson.dependencies) {
      projectDeps.push(...Object.keys(project.packageJson.dependencies));
    }
    if (project.packageJson.devDependencies) {
      projectDevDeps.push(...Object.keys(project.packageJson.devDependencies));
    }

    const allDeps = projectDeps.concat(projectDevDeps);
    if (allDeps.length === 0) {
      return;
    }

    if (!allDeps.includes('sp-pnp-js') ||
      !allDeps.includes('@pnp/sp')) {
      return;
    }

    const dependencyType = projectDeps.includes('sp-pnp-js') ? 'dependencies' : 'devDependencies';
    const node = this.getAstNodeFromFile(project.packageJson as PackageJson, `${dependencyType}.sp-pnp-js`);
    this.addFindingWithCustomInfo(
      this.title,
      this.description,
      [{
        file: this.file,
        resolution: `${dependencyType === 'dependencies' ? 'uninstall' : 'uninstallDev'} sp-pnp-js`,
        position: this.getPositionFromNode(node)
      }],
      findings
    );
  }
}