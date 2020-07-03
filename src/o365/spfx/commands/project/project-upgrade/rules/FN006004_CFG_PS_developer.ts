import { Finding } from "../";
import { Project } from "../../model";
import { Rule } from "./Rule";

export class FN006004_CFG_PS_developer extends Rule {
  constructor() {
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
  };

  get resolution(): string {
    return `{
  "solution": {
    "developer": {
      "name": "Contoso",
      "privacyUrl": "https://contoso.com/privacy",
      "termsOfUseUrl": "https://contoso.com/terms-of-use",
      "websiteUrl": "https://contoso.com/my-app",
      "mpnId": "000000"
    }
  }
}`;
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Optional';
  };

  get file(): string {
    return './config/package-solution.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageSolutionJson ||
      !project.packageSolutionJson.solution) {
      return;
    }

    if (!project.packageSolutionJson.solution.developer) {
      this.addFinding(findings);
    }
  }
}