import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011001_MAN_webpart_schema extends ManifestRule {
  constructor(private schema: string) {
    super()/* istanbul ignore next */;
  }

  get id(): string {
    return 'FN011001';
  }

  get resolution(): string {
    return `{
  "$schema": "${this.schema}"
}`;
  };

  get severity(): string {
    return 'Required';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length === 0) {
      return;
    }

    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.$schema !== this.schema) {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Web part manifest schema', `Update schema in manifest ${relativePath}`, this.resolution, relativePath, findings);
      }
    });
  }
}