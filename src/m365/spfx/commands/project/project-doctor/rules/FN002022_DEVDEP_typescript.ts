import { satisfies, minVersion } from 'semver';
import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN002022_DEVDEP_typescript extends JsonRule {
  constructor(private version: string) {
    super();
  }
  get id(): string {
    return 'FN002022';
  }
  get title(): string {
    return 'TypeScript version';
  }
  get description(): string {
    return `TypeScript version in package.json should be at least @${this.version}`;
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
    if (!project.packageJson) {
      return;
    }
    const tsVersion = project.packageJson.devDependencies?.typescript || project.packageJson.dependencies?.typescript;
    if (!tsVersion) {
      const node = this.getAstNodeFromFile(project.packageJson, 'devDependencies') ||
        this.getAstNodeFromFile(project.packageJson, 'dependencies');
      this.addFindingWithCustomInfo(`TypeScript is not specified`, `Add TypeScript devDependency at least @${this.version}`,
        [{
          file: this.file,
          resolution: `installDev typescript@${this.version}`,
          position: this.getPositionFromNode(node)
        }],
        findings
      );
      return;
    }

    const minTsVersion = minVersion(tsVersion);
    if (minTsVersion && satisfies(minTsVersion.version, this.version)) {
      return;
    }
    const node = this.getAstNodeFromFile(project.packageJson, 'devDependencies.typescript') ||
      this.getAstNodeFromFile(project.packageJson, 'dependencies.typescript');
    this.addFindingWithCustomInfo(`TypeScript version is lower than @${this.version}`, `Update TypeScript to at least @${this.version}`,
      [{
        file: this.file,
        resolution: `installDev typescript@${this.version}`,
        position: this.getPositionFromNode(node)
      }],
      findings
    );
  }
}