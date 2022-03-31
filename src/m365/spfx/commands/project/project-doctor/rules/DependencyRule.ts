import { coerce, satisfies, SemVer } from 'semver';
import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export abstract class DependencyRule extends JsonRule {
  constructor(protected packageName: string, protected supportedRange: string, protected isDevDep: boolean = false) {
    super();
  }

  get title(): string {
    return this.packageName;
  }

  get description(): string {
    return '';
  }

  get resolution(): string {
    return `${(this.isDevDep ? 'installDev' : 'install')} ${this.packageName}@${this.supportedRange.includes(' ') ? `"${this.supportedRange}"` : this.supportedRange}`;
  }

  get resolutionType(): string {
    return 'cmd';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  customCondition(project: Project): boolean {
    return true;
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson || !this.customCondition(project)) {
      return;
    }

    // if the project has no dependencies, then we assume the package is missing
    let packageNotFound = (this.isDevDep && !project.packageJson.devDependencies) ||
      (!this.isDevDep && !project.packageJson.dependencies);
    let packageVersionFromProject: string | undefined = undefined;
    let minSemVer: SemVer | null = null;

    if (!packageNotFound) {
      // try to get the current version of the dependency installed in the
      // project. If not possible, we assume the dependency is missing
      packageVersionFromProject = this.isDevDep ? project.packageJson.devDependencies![this.packageName] : project.packageJson.dependencies![this.packageName];
      if (!packageVersionFromProject) {
        packageNotFound = true;
      }
      else {
        minSemVer = coerce(packageVersionFromProject);
        if (!minSemVer) {
          packageNotFound = true;
        }
      }
    }

    if (packageNotFound) {
      return this.addFindingWithCustomInfo(
        this.title,
        `Install missing package ${this.packageName}`,
        [{
          file: this.file,
          resolution: this.resolution
        }],
        findings);
    }

    if (satisfies(minSemVer!, this.supportedRange)) {
      return;
    }

    const node = this.getAstNodeFromFile(project.packageJson, `${(this.isDevDep ? 'devDependencies' : 'dependencies')}.${this.packageName}`);
    this.addFindingWithCustomInfo(
      this.title,
      `Install supported version of the ${this.packageName} package`,
      [{
        file: this.file,
        resolution: this.resolution,
        position: this.getPositionFromNode(node)
      }],
      findings);
  }
}