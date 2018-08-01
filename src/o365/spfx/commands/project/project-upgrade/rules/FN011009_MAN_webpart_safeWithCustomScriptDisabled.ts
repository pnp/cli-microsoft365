import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011009_MAN_webpart_safeWithCustomScriptDisabled extends ManifestRule {
  constructor(private add: boolean) {
    super()/* istanbul ignore next */;
  }

  get id(): string {
    return 'FN011009';
  }

  get resolution(): string {
    return `{
  "safeWithCustomScriptDisabled": false
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
      if (manifest.componentType !== 'WebPart') {
        return;
      }

      if ((this.add && manifest.safeWithCustomScriptDisabled === undefined) ||
        (!this.add && manifest.safeWithCustomScriptDisabled !== undefined)) {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Web part manifest safeWithCustomScriptDisabled', `${this.add ? 'Update' : 'Remove'} the safeWithCustomScriptDisabled property in manifest ${relativePath}`, this.resolution, relativePath, findings);
      }
    });
  }
}