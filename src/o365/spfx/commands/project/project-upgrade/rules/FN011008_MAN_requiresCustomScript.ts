import { Finding } from '../Finding';
import { Project } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011008_MAN_requiresCustomScript extends ManifestRule {
  get id(): string {
    return 'FN011008';
  }

  get resolution(): string {
    return '';
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
      if (typeof manifest.safeWithCustomScriptDisabled !== 'undefined') {
        const relativePath: string = path.relative(project.path, manifest.path);
        this.addFindingWithCustomInfo('Client-side component manifest requiresCustomScript property', `In manifest ${relativePath} rename the safeWithCustomScriptDisabled property to requiresCustomScript and invert its value`, `{
  "requiresCustomScript": ${!manifest.safeWithCustomScriptDisabled}
`, relativePath, findings);
      }
    });
  }
}