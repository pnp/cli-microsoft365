import { Finding, Occurrence } from '../../report-model/index.js';
import { Project } from '../../project-model/index.js';
import { ManifestRule } from "./ManifestRule.js";

export class FN011009_MAN_webpart_safeWithCustomScriptDisabled extends ManifestRule {
  private add: boolean;

  constructor(options: { add: boolean }) {
    super();
    this.add = options.add;
  }

  get id(): string {
    return 'FN011009';
  }

  get title(): string {
    return 'Web part manifest safeWithCustomScriptDisabled';
  }

  get description(): string {
    return `${this.add ? 'Update' : 'Remove'} the safeWithCustomScriptDisabled property in the manifest`;
  }

  get resolution(): string {
    return `{
  "safeWithCustomScriptDisabled": false
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
      if (manifest.componentType !== 'WebPart') {
        return;
      }

      if ((this.add && manifest.safeWithCustomScriptDisabled === undefined) ||
        (!this.add && manifest.safeWithCustomScriptDisabled !== undefined)) {
        const node = this.getAstNodeFromFile(manifest, 'safeWithCustomScriptDisabled');
        this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
      }
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}