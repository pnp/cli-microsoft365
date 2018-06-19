import { Finding } from "../";
import { Project, CommandSetManifest } from "../model";
import * as path from 'path';
import { ManifestRule } from "./ManifestRule";

export class FN011005_MAN_listViewCommandSet_items extends ManifestRule {

  get id(): string {
    return 'FN011005';
  }

  get title(): string {
    return `List view command set manifest items`;
  }

  get description(): string {
    return `Replace the "commands" property with "items" property`;
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

        let items: string = '';
        const commands: any = (commandSetManifest as any)["commands"];

        if (commands !== undefined) {

          // convert the "commands" schema to "items" schema
          Object.keys(commands).forEach(key => {
            const valueObj = commands[key];

            let props: string = '';
            Object.keys(valueObj).forEach(prop => {

              if (prop !== "title") {

                props += `"${prop}": "${valueObj[prop]}",
              `;
              }
            });

            // remove ending ','
            props = props.substring(0, props.lastIndexOf(','));

            // add type if missing
            if (valueObj.type === undefined) {
              props += `,
                "type": "command"`;
            }

            items += `"${key}": {
                "title": { "default": "${valueObj.title}" },
                ${props}
              },
              `;
          });

          // remove ending ','
          items = items.substring(0, items.lastIndexOf(','));

          const resolution: string = `{
            "items": {
              ${items}
            } 
          }`;

          this.addFindingWithCustomInfo(this.title, `${this.description} in file ${relativePath}`, resolution, relativePath, findings);
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