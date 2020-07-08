import { Finding, Occurrence } from '../';
import { Project } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011008_MAN_requiresCustomScript extends ManifestRule {
  get id(): string {
    return 'FN011008';
  }

  get supersedes(): string[] {
    return ['FN011009'];
  }

  get title(): string {
    return 'Client-side component manifest requiresCustomScript property';
  }

  get description(): string {
    return 'In the manifest rename the safeWithCustomScriptDisabled property to requiresCustomScript and invert its value';
  };

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

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      if (typeof manifest.safeWithCustomScriptDisabled !== 'undefined') {
        this.addOccurrence(`{
  "requiresCustomScript": ${!manifest.safeWithCustomScriptDisabled}
}`, manifest.path, project.path, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
