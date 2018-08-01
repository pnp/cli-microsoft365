import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011010_MAN_webpart_version extends ManifestRule {
  get id(): string {
    return 'FN011010';
  }

  get resolution(): string {
    return `{
      "version": "*",
}`;
  };

  get severity(): string {
    return 'Optional';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length === 0) {
      return;
    }

    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.version !== '*') {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Web part manifest version', `Update version in manifest ${relativePath} to use automated component versioning`, this.resolution, relativePath, findings);
      }
    });
  }
}