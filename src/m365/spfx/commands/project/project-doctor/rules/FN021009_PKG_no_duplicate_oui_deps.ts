import { coerce, SemVer } from 'semver';
import { JsonRule } from '../../JsonRule';
import { PackageJson, Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN021009_PKG_no_duplicate_oui_deps extends JsonRule {
  get id(): string {
    return 'FN021009';
  }

  get title(): string {
    return '@fluentui/react and office-ui-fabric-react installed in the project';
  }

  get description(): string {
    return '@fluentui/react and office-ui-fabric-react installed in the project. Consider uninstalling @fluentui/react';
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

  customCondition(project: Project): boolean {
    const ouifSemVer = coerce(project.packageJson?.dependencies?.['office-ui-fabric-react']);
    if (!ouifSemVer) {
      return false;
    }

    // ouif and @fluentui/react are both required starting from 7.199.x
    return ouifSemVer < new SemVer('7.199.0');
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson || !this.customCondition(project)) {
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
    if (!allDeps.includes('office-ui-fabric-react') ||
      !allDeps.includes('@fluentui/react')) {
      return;
    }

    const dependencyType = projectDeps.includes('@fluentui/react') ? 'dependencies' : 'devDependencies';
    const node = this.getAstNodeFromFile(project.packageJson as PackageJson, `${dependencyType}.@fluentui/react`);
    this.addFindingWithCustomInfo(
      this.title,
      this.description,
      [{
        file: this.file,
        resolution: `${dependencyType === 'dependencies' ? 'uninstall' : 'uninstallDev'} @fluentui/react`,
        position: this.getPositionFromNode(node)
      }],
      findings
    );
  }
}