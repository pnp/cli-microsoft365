import { Finding } from "../";
import { Project, CommandSetManifest } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011006_MAN_listViewCommandSet_items extends ManifestRule {
  get id(): string {
    return 'FN011006';
  }

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

      const relativePath: string = path.relative(project.path, manifest.path);
      this.addFindingWithCustomInfo('List view command set items property', `In the manifest ${relativePath} add the items property`, JSON.stringify(resolution, null, 2), relativePath, findings);
    });
  }
}