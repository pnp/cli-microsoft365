import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export class FN012021_TSC_excludeDirectories extends JsonRule {
  private directories: string[];

  constructor(options: { directories: string[] }) {
    super();
    this.directories = options.directories;
  }

  get id(): string {
    return 'FN012021';
  }

  get title(): string {
    return 'tsconfig.json watchOptions.excludeDirectories property';
  }

  get description(): string {
    return `Update tsconfig.json watchOptions.excludeDirectories property`;
  }

  get resolution(): string {
    return `{
  "watchOptions": {
    "excludeDirectories": ${JSON.stringify(this.directories)}
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './tsconfig.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsConfigJson) {
      return;
    }

    if (!project.tsConfigJson.watchOptions?.excludeDirectories ||
      JSON.stringify(project.tsConfigJson.watchOptions.excludeDirectories) !== JSON.stringify(this.directories)) {
      const node = this.getAstNodeFromFile(project.tsConfigJson, 'watchOptions.excludeDirectories');
      this.addFindingWithPosition(findings, node);
    }
  }
}