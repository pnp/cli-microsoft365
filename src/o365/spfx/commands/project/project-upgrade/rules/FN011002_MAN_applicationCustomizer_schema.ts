import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011002_MAN_applicationCustomizer_schema extends ManifestRule {
  constructor(private schema: string) {
    super()/* istanbul ignore next */;
  }

  get id(): string {
    return 'FN011002';
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
      if (manifest.componentType === 'Extension' &&
        manifest.extensionType === 'ApplicationCustomizer' &&
        manifest.$schema !== this.schema) {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Application customizer manifest schema', `Update schema in manifest ${relativePath}`, this.resolution, relativePath, findings);
      }
    });
  }
}