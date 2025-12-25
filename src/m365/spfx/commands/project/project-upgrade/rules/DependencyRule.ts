import semver from 'semver';
import { Hash } from '../../../../../../utils/types.js';
import { JsonRule } from '../../JsonRule.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';

export interface DependencyRuleOptions {
  packageName: string;
  packageVersion: string;
  isDevDep?: boolean;
  isOptional?: boolean;
  add?: boolean;
}

export abstract class DependencyRule extends JsonRule {
  protected packageName: string;
  protected packageVersion: string;
  protected isDevDep: boolean;
  protected isOptional: boolean;
  protected add: boolean;

  constructor(options: DependencyRuleOptions) {
    super();
    const { packageName, packageVersion, isDevDep = false, isOptional = false, add = true } = options;
    this.packageName = packageName;
    this.packageVersion = packageVersion;
    this.isDevDep = isDevDep;
    this.isOptional = isOptional;
    this.add = add;
  }

  get title(): string {
    return this.packageName;
  }

  get description(): string {
    return `${(this.add ? 'Upgrade' : 'Remove')} SharePoint Framework ${(this.isDevDep ? 'dev ' : '')}dependency package ${this.packageName}`;
  }

  get resolution(): string {
    return this.add ?
      `${(this.isDevDep ? 'installDev' : 'install')} ${this.packageName}@${this.packageVersion}` :
      `${(this.isDevDep ? 'uninstallDev' : 'uninstall')} ${this.packageName}`;
  }

  get resolutionType(): string {
    return 'cmd';
  }

  get severity(): string {
    return 'Required';
  }

  get file(): string {
    return './package.json';
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  customCondition(project: Project): boolean {
    return false;
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.packageJson) {
      return;
    }

    const projectDependencies: Hash | undefined = this.isDevDep ? project.packageJson.devDependencies : project.packageJson.dependencies;
    const versionEntry: string | null = projectDependencies ? projectDependencies[this.packageName] : '';
    if (this.add) {
      let jsonProperty: string = this.isDevDep ? 'devDependencies' : 'dependencies';

      if (versionEntry) {
        jsonProperty += `.${this.packageName}`;

        if (this.#needsUpdate(this.packageVersion, versionEntry)) {
          const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
          this.addFindingWithPosition(findings, node);
        }
      }
      else {
        if (!this.isOptional || this.customCondition(project)) {
          const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
          this.addFindingWithCustomInfo(this.packageName, this.description.replace('Upgrade', 'Install'), [{
            file: this.file,
            resolution: this.resolution,
            position: this.getPositionFromNode(node)
          }], findings);
        }
      }
    }
    else {
      const jsonProperty: string = `${(this.isDevDep ? 'devDependencies' : 'dependencies')}.${this.packageName}`;

      if (versionEntry) {
        const node = this.getAstNodeFromFile(project.packageJson, jsonProperty);
        this.addFindingWithPosition(findings, node);
      }
    }
  }

  /**
 * Determines if a package needs to be updated based on a rule version
 * @param {string} ruleVersion - The version/range from the rule (e.g., '5.8.1', '~5.8.0', '^6.0.0')
 * @param {string} currentVersion - The version/range from package.json
 * @returns {boolean} - true if update is needed
 */
  #needsUpdate(ruleVersion: string, currentVersion: string): boolean {
    try {
      // Get minimum versions for both
      const ruleMin = semver.minVersion(ruleVersion);
      const currentMin = semver.minVersion(currentVersion);

      // Check if ranges overlap
      const rangesOverlap = semver.intersects(ruleVersion, currentVersion);

      if (rangesOverlap) {
        // Even if they overlap, update if rule requires a higher minimum version
        if (ruleMin && currentMin && semver.gt(ruleMin, currentMin)) {
          return true;
        }
        return false;
      }

      // Ranges don't overlap - check if rule range is greater
      // Get the maximum version that satisfies the current range
      const currentMax = this.#getMaxVersion(currentVersion);

      // If rule's minimum is greater than current's maximum, update is needed
      return !!(ruleMin && currentMax && semver.gt(ruleMin, currentMax));
    }
    catch {
      return false;
    }
  }

  /**
   * Gets the maximum version from a range
   * For open-ended ranges like '>=1.0.0', returns the minVersion
   * For bounded ranges, returns the upper bound
   */
  #getMaxVersion(range: string): semver.SemVer | null {
    const rangeObj = new semver.Range(range);

    // If it's a specific version (no range operators), return it
    if (semver.valid(range)) {
      return semver.parse(range);
    }

    // For ranges, get the highest version from the set
    // Check the range set to find upper bounds
    let maxVer = null;

    for (const comparatorSet of rangeObj.set) {
      for (const comparator of comparatorSet) {
        if (comparator.operator === '<' || comparator.operator === '<=') {
          const ver = comparator.semver;
          if (!maxVer || semver.gt(ver, maxVer)) {
            maxVer = ver;
          }
        }
      }
    }

    // If no upper bound found, use minVersion as fallback
    return maxVer || semver.minVersion(range);
  }
}