import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011001_MAN_webpart_schema extends ManifestRule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN011001';
  }

  get title(): string {
    return 'Web part manifest schema';
  }

  get description(): string {
    return 'Update schema in manifest';
  };

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

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.$schema !== this.schema) {
        this.addOccurrence(this.resolution, manifest.path, project.path, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}