import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011005_MAN_webpart_defaultGroup extends ManifestRule {
  constructor(private oldDefaultGroup: string, private newDefaultGroup: string) {
    super()/* istanbul ignore next */;
  }

  get id(): string {
    return 'FN011005';
  }

  get resolution(): string {
    return `{
  "preconfiguredEntries": [{
    "group": { "default": "${this.newDefaultGroup}" }
  }]
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
        manifest.preconfiguredEntries) {
        manifest.preconfiguredEntries.forEach(e => {
          if (e.group && e.group.default === this.oldDefaultGroup) {
            const relativePath: string = path.relative(project.path, manifest.path);
            this.addFindingWithCustomInfo('Web part manifest default group', `In the manifest ${relativePath} update the default group value`, this.resolution, relativePath, findings);
          }
        });
      }
    });
  }
}