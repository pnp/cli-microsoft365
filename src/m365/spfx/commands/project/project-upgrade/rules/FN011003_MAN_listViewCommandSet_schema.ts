import { Finding, Occurrence } from '../../report-model';
import { Project } from '../../project-model';
import { ManifestRule } from "./ManifestRule";

export class FN011003_MAN_listViewCommandSet_schema extends ManifestRule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN011003';
  }

  get title(): string {
    return 'List view command set manifest schema';
  }

  get description(): string {
    return 'Update schema in manifest';
  }

  get resolution(): string {
    return `{
  "$schema": "${this.schema}"
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
      if (manifest.componentType === 'Extension' &&
        manifest.extensionType === 'ListViewCommandSet' &&
        manifest.$schema !== this.schema) {
        const node = this.getAstNodeFromFile(manifest, '$schema');
        this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}