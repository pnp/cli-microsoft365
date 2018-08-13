import { Finding } from "../";
import { Project, CommandSetManifest } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011006_MAN_listViewCommandSet_items extends ManifestRule {

  get id(): string {
    return 'FN011006';
  }

  get title(): string {
    return `List view command set manifest items`;
  }

  get description(): string {
    return `Add "items" property (to replace the "commands" property)`;
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

    project.manifests.forEach(manifest => {
      const commandSetManifest: CommandSetManifest = manifest as CommandSetManifest;

      if (commandSetManifest.componentType === 'Extension' &&
        commandSetManifest.extensionType === 'ListViewCommandSet' &&
        commandSetManifest.items === undefined) {

        const relativePath: string = path.relative(project.path, manifest.path);

        let resolution: any = { items: {} };
        const commands: any = (commandSetManifest as any)["commands"];

        if (commands !== undefined) {

          // convert the "commands" schema to "items" schema
          Object.keys(commands).forEach(key => {
            const valueObj = commands[key];

            resolution.items[key] = { title: { default: valueObj.title }};

            Object.keys(valueObj).forEach(prop => {

              if (prop !== "title") {
                resolution.items[key][prop] = valueObj[prop]
              }
            });

            // add type if missing
            if (valueObj.type === undefined) {
              resolution.items[key].type = "command";
            }
          });

          this.addFindingWithCustomInfo(this.title, `${this.description} in file ${relativePath}`, JSON.stringify(resolution, null, 2), relativePath, findings);
        } else {
          // this should not happen
          // if no items prop, but also no commands prop
          // we cannot do other than notify that
          // the config requires items setup
          const resolution: string = `The "items" property is missing in ${relativePath}. Please it setup.`;
          this.addFindingWithCustomInfo(this.title, `${resolution}. Visit the schema url for more inforamtion.`, resolution, this.file, findings);
        }
      }
    });
  }
}