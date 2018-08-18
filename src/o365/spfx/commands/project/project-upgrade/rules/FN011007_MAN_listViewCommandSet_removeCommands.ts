import { Finding } from "../";
import { Project, CommandSetManifest } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011007_MAN_listViewCommandSet_removeCommands extends ManifestRule {
  get id(): string {
    return 'FN011007';
  }

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

    project.manifests.forEach(manifest => {
      const commandSetManifest: CommandSetManifest = manifest as CommandSetManifest;

      if (commandSetManifest.componentType !== 'Extension' ||
        commandSetManifest.extensionType !== 'ListViewCommandSet' ||
        !commandSetManifest.commands) {
        return;
      }

      const relativePath: string = path.relative(project.path, manifest.path);
      this.addFindingWithCustomInfo('List view command set commands property', `In the manifest ${relativePath} remove the commands property`, JSON.stringify({ commands: commandSetManifest.commands }, null, 2), relativePath, findings);
    });
  }
}