import { Finding, Occurrence } from '../../report-model/index.js';
import { Project } from '../../project-model/index.js';
import { ManifestRule } from "./ManifestRule.js";

export class FN011004_MAN_fieldCustomizer_schema extends ManifestRule {
  private schema: string;
  constructor(options: { schema: string }) {
    super();
    this.schema = options.schema;
  }

  get id(): string {
    return 'FN011004';
  }

  get title(): string {
    return 'Field customizer manifest schema';
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
        manifest.extensionType === 'FieldCustomizer' &&
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