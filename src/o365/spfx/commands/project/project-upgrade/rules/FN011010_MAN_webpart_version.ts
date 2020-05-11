import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011010_MAN_webpart_version extends ManifestRule {
  get id(): string {
    return 'FN011010';
  }

  get title(): string {
    return 'Web part manifest version';
  }

  get description(): string {
    return 'Update version in manifest to use automated component versioning';
  };

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

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.version !== '*') {
        this.addOccurrence(this.resolution, manifest.path, project.path, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
