import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011011_MAN_webpart_supportedHosts extends ManifestRule {
  constructor(private add: boolean) {
    super();
  }

  get id(): string {
    return 'FN011011';
  }

  get title(): string {
    return 'Web part manifest supportedHosts';
  }

  get description(): string {
    return `${this.add ? 'Update' : 'Remove'} the supportedHosts property in the manifest`;
  };

  get resolution(): string {
    return `{
  "supportedHosts": ["SharePointWebPart"]
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
      if (manifest.componentType !== 'WebPart') {
        return;
      }

      if ((this.add && manifest.supportedHosts === undefined) ||
        (!this.add && manifest.supportedHosts !== undefined)) {
        this.addOccurrence(this.resolution, manifest.path, project.path, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}