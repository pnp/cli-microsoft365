import { Finding, Occurrence } from "../";
import { Project, CommandSetManifest } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011006_MAN_listViewCommandSet_items extends ManifestRule {
  get id(): string {
    return 'FN011006';
  }

  get title(): string {
    return 'List view command set items property';
  }

  get description(): string {
    return 'In the manifest add the items property';
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
      const commandSetManifest: CommandSetManifest = manifest as CommandSetManifest;

      if (commandSetManifest.componentType !== 'Extension' ||
        commandSetManifest.extensionType !== 'ListViewCommandSet' ||
        commandSetManifest.items) {
        return;
      }

      const resolution: any = {
        items: commandSetManifest.commands || {}
      };
      Object.keys(resolution.items).forEach(k => {
        resolution.items[k].title = { default: resolution.items[k].title },
          resolution.items[k].type = 'command';
      });

      this.addOccurrence(JSON.stringify(resolution, null, 2), manifest.path, project.path, occurrences);
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}