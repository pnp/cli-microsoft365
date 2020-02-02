import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011005_MAN_webpart_defaultGroup extends ManifestRule {
  constructor(private oldDefaultGroup: string, private newDefaultGroup: string) {
    super();
  }

  get id(): string {
    return 'FN011005';
  }

  get title(): string {
    return 'Web part manifest default group';
  }

  get description(): string {
    return 'In the manifest update the default group value';
  };

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

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.preconfiguredEntries) {
        manifest.preconfiguredEntries.forEach(e => {
          if (e.group && e.group.default === this.oldDefaultGroup) {
            this.addOccurrence(this.resolution, manifest.path, project.path, occurrences);
          }
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}