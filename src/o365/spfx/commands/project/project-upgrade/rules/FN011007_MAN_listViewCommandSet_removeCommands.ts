import { Finding, Occurrence } from "../";
import { Project, CommandSetManifest } from "../../model";
import { ManifestRule } from "./ManifestRule";

export class FN011007_MAN_listViewCommandSet_removeCommands extends ManifestRule {
  get id(): string {
    return 'FN011007';
  }

  get title(): string {
    return 'List view command set commands property';
  }

  get description(): string {
    return 'In the manifest remove the commands property';
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
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
        !commandSetManifest.commands) {
        return;
      }

      this.addOccurrence(JSON.stringify({ commands: commandSetManifest.commands }, null, 2), manifest.path, project.path, occurrences);
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}