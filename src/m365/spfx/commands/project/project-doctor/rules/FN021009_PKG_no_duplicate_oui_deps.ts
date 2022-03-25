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