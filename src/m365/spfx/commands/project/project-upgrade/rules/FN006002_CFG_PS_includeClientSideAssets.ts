import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN006002_CFG_PS_includeClientSideAssets extends JsonRule {
  constructor(private includeClientSideAssets: boolean) {
    super();
  }

  get id(): string {
    return 'FN006002';
  }

  get title(): string {
    return 'package-solution.json includeClientSideAssets';
  }

  get description(): string {
    return `Update package-solution.json includeClientSideAssets`;
  };

  get resolution(): string {
    return `{
  "solution": {
    "includeClientSideAssets": ${this.includeClientSideAssets}
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Required';
  };

  get file(): string {
    return './config/package-solution.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    if (project.packageSolutionJson.solution.includeClientSideAssets !== this.includeClientSideAssets) {
      const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution.includeClientSideAssets');
      this.addFindingWithPosition(findings, node);
    }
  }
}