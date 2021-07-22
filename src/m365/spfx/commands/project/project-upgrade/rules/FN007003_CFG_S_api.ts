import { Finding } from "../";
import { Project } from "../../model";
import { JsonRule } from "./JsonRule";

export class FN007003_CFG_S_api extends JsonRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN007003';
  }

  get title(): string {
    return 'serve.json api';
  }

  get description(): string {
    return `From serve.json remove the api property`;
  }

  get resolution(): string {
    return ``;
  }

  get resolutionType(): string {
    return 'json';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './config/serve.json';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.serveJson) {
      return;
    }

    if (project.serveJson.api) {
      const node = this.getAstNodeFromFile(project.serveJson, 'api');
      this.addFindingWithPosition(findings, node);
    }
  }
}