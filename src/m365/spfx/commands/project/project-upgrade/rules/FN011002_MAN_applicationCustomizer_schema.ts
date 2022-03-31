import { Finding, Occurrence } from '../../report-model';
import { Project } from '../../project-model';
import { ManifestRule } from "./ManifestRule";

export class FN011002_MAN_applicationCustomizer_schema extends ManifestRule {
  constructor(private schema: string) {
    super();
  }

  get id(): string {
    return 'FN011002';
  }

  get title(): string {
    return 'Application customizer manifest schema';
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
        manifest.extensionType === 'ApplicationCustomizer' &&
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