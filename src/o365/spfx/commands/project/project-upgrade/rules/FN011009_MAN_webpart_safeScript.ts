import { Finding } from "../";
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011009_MAN_webpart_safeScript extends ManifestRule {
  get id(): string {
    return 'FN011009';
  }

  get resolution(): string {
    return `{
      /**
       * This property should only be set to true if it is certain that the webpart does not
       *  allow arbitrary scripts to be called
       */
      "safeWithCustomScriptDisabled": false,
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
        manifest.safeWithCustomScriptDisabled === undefined) {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Web part manifest safeWithCustomScriptDisabled', `Update safeWithCustomScriptDisabled in manifest ${relativePath}`, this.resolution, relativePath, findings);
      }
    });
  }
}