import { Finding } from "../";
import { Project, CommandSetManifest } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011007_MAN_listViewCommandSet_commands extends ManifestRule {

  get id(): string {
    return 'FN011007';
  }

  get title(): string {
    return `List view command set manifest commands`;
  }

  get description(): string {
    return `Remove the "commands" property`;
  };

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'json';
  };

  get severity(): string {
    return 'Recomended';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length === 0) {
      return;
    }

    project.manifests.forEach(manifest => {
      const commandSetManifest: CommandSetManifest = manifest as CommandSetManifest;

      if (commandSetManifest.componentType === 'Extension' &&
        commandSetManifest.extensionType === 'ListViewCommandSet' &&
        commandSetManifest.items === undefined) {

        const relativePath: string = path.relative(project.path, manifest.path);
        const commands: any = (commandSetManifest as any)["commands"];

        if (commands !== undefined) {

          this.addFindingWithCustomInfo(this.title, `${this.description} in file ${relativePath}`, JSON.stringify({ commands: commands }, null, 2), relativePath, findings);
        }
      }
    });
  }
}