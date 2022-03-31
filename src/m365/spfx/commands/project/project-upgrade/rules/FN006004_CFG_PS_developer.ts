import { JsonRule } from '../../JsonRule';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';

export class FN006004_CFG_PS_developer extends JsonRule {
  constructor(private version?: string) {
    super();
  }

  get id(): string {
    return 'FN006004';
  }

  get title(): string {
    return 'package-solution.json developer';
  }

  get description(): string {
    return `In package-solution.json add developer section`;
  }

  get resolution(): string {
    return `{
  "solution": {
    "developer": {
      "name": "",
      "privacyUrl": "",
      "termsOfUseUrl": "",
      "websiteUrl": "",
      "mpnId": "${this.version ? `Undefined-${this.version}` : ''}"
    }
  }
}`;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Optional';
  }

  get file(): string {
    return './config/package-solution.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    if (!project.packageSolutionJson.solution.developer) {
      const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution');
      this.addFindingWithPosition(findings, node);
    }
  }
}