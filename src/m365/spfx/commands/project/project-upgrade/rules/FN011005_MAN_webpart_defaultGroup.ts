import { Finding, Occurrence } from '../../report-model';
import { Project } from '../../project-model';
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
  }

  get resolution(): string {
    return `{
  "preconfiguredEntries": [{
    "group": { "default": "${this.newDefaultGroup}" }
  }]
}`;
  }

  get severity(): string {
    return 'Required';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length === 0) {
      return;
    }

    const occurrences: Occurrence[] = [];
    project.manifests.forEach(manifest => {
      if (manifest.componentType === 'WebPart' &&
        manifest.preconfiguredEntries) {
        manifest.preconfiguredEntries.forEach((e, i) => {
          if (e.group && e.group.default === this.oldDefaultGroup) {
            const node = this.getAstNodeFromFile(manifest, `preconfiguredEntries[${i}].group.default`);
            this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
          }
        });
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}