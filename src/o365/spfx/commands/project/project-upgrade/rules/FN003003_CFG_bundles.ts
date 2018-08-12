import { Finding } from "../";
import { Project } from "../model";
import { Rule } from "./Rule";

export class FN003003_CFG_bundles extends Rule {

  get id(): string {
    return 'FN003003';
  }

  get title(): string {
    return `config.json bundles`;
  }

  get description(): string {
    return `Add "bundles" property (to replace the "entries" property) in ${this.file}`;
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

  get file(): string {
    return './config/config.json';
  };

  visit(project: Project, findings: Finding[]): void {
    if (!project.configJson) {
      return;
    }

    if (project.configJson.bundles === undefined) {

      const entries: any = (project.configJson as any).entries;

      if (entries !== undefined) {

        let resolution: any = { bundles: {} };

        // convert the "entries" schema to "bundles" schema
        Object.keys(entries).forEach(key => {
          const valueObj = entries[key];
          let bundleName: string = key;
          let props: any = {};

          Object.keys(valueObj).forEach(prop => {

            switch (prop) {
              case "entry":
                props.entrypoint = valueObj[prop];
                bundleName = this.tryGetBundleName(bundleName, valueObj[prop]);
                break;
              case "outputPath":
                // skip. Should not be in the file anymore
                break;
              default:
                props[prop] = valueObj[prop];
            }
          });
          resolution.bundles[bundleName] = {components:[props]};
        });

        this.addFindingWithCustomInfo(this.title, this.description, JSON.stringify(resolution, null, 2), this.file, findings);
      }
      else {
        // this should not happen
        // if no bundles prop, but also no entries prop
        // we cannot do other than notify that
        // the config requires bundles setup
        const resolution: string = `The "bundles" property is missing in ${this.file}. Please it setup.`;
        this.addFindingWithCustomInfo(this.title, `${resolution}. Visit the schema url for more inforamtion.`, resolution, this.file, findings);
      }
    }
  }

  /**
   * Smart guess on the bundle name.
   * @param bundleName any existing bundle name
   * @param entrypointValue the entrypoint value
   */
  private tryGetBundleName(bundleName: string, entrypointValue: string): string {

    let name: string = '';
    try {
      name = entrypointValue.substring(entrypointValue.lastIndexOf('/') + 1, entrypointValue.length);
      name = name.replace('.js', '');
      name = name.replace(/([a-z](?=[A-Z]))/g, '$1-');
      bundleName = name.toLowerCase();
    } catch {
      // if it cannot smart guess, leave the existing name (index)
    }

    return bundleName;
  }

}