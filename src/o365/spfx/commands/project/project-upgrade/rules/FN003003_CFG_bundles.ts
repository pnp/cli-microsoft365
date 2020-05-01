import { Finding } from "../";
import { Project, Entry } from "../../model";
import { Rule } from "./Rule";

export class FN003003_CFG_bundles extends Rule {
  get id(): string {
    return 'FN003003';
  }

  get title(): string {
    return `config.json bundles`;
  }

  get description(): string {
    return `In config.json add the 'bundles' property`;
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
    if (!project.configJson ||
      typeof project.configJson.bundles !== 'undefined') {
      return;
    }

    const entries: Entry[] | undefined = project.configJson.entries;
    if (!entries) {
      return;
    }

    const resolution: any = { bundles: {} };
    entries.forEach((e, i) => {
      resolution.bundles[this.tryGetBundleName(e.entry, i)] = {
        components: [{
          entrypoint: e.entry,
          manifest: e.manifest
        }]
      };
    });

    this.addFindingWithOccurrences([{
      file: this.file,
      resolution: JSON.stringify(resolution, null, 2)
    }], findings);
  }

  /**
   * Smart guess on the bundle name.
   * @param entry the entry value
   */
  private tryGetBundleName(entry: string, index: number): string {
    let name: string = index.toString();
    name = entry.substring(entry.lastIndexOf('/') + 1, entry.length);
    name = name.replace('.js', '');
    name = name.replace(/([a-z](?=[A-Z]))/g, '$1-');
    name = name.toLowerCase();

    return name;
  }
}