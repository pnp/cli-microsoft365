import { Range, coerce, gt, validRange } from 'semver';
import { Project } from '../m365/spfx/commands/project/project-model/index.js';

export const spfx = {
  isReactProject(project: Project): boolean {
    return (typeof project.yoRcJson !== 'undefined' &&
      typeof project.yoRcJson['@microsoft/generator-sharepoint'] !== 'undefined' &&
      (project.yoRcJson["@microsoft/generator-sharepoint"].framework === 'react' ||
        project.yoRcJson["@microsoft/generator-sharepoint"].template === 'react')) ||
      typeof project.packageJson?.dependencies?.['react'] !== 'undefined';
  },

  isKnockoutProject(project: Project): boolean {
    return (typeof project.yoRcJson !== 'undefined' &&
      typeof project.yoRcJson['@microsoft/generator-sharepoint'] !== 'undefined' &&
      project.yoRcJson["@microsoft/generator-sharepoint"].framework === 'knockout') ||
      typeof project.packageJson?.dependencies?.['knockout'] !== 'undefined';
  },

  getHighestNodeVersion(versionRange: string): string {
    if (!versionRange) {
      throw new Error('Node version range was not provided.');
    }

    const ranges = versionRange
      .split('||')
      .map(range => range.trim())
      .filter(range => range.length > 0);

    let highestMajor: number | null = null;
    let exactVersion: string | null = null;

    const simpleVersionPattern = /^\d+(\.\d+(\.\d+)?)?$/;
    if (ranges.every(r => simpleVersionPattern.test(r))) {
      const highest = ranges.reduce((best, curr) =>
        gt(coerce(curr)!, coerce(best)!) ? curr : best
      );
      const parts = highest.split('.');
      if (parts.length >= 3) {
        return highest;
      }

      if (parts.length === 2) {
        return `${highest}.x`;
      }
    }

    for (const rangeString of ranges) {
      const normalized = validRange(rangeString);
      if (!normalized) {
        continue;
      }

      const rangeObj = new Range(normalized);
      let maxMajor = 0;

      // Analyze the range to find the maximum major version
      for (const comparatorSet of rangeObj.set) {
        for (const comparator of comparatorSet) {
          if (comparator.operator === '<') {
            // Exclusive upper bound: <17.0.0 means max major is 16
            maxMajor = Math.max(maxMajor, comparator.semver.major - 1);
          }
          else if (comparator.operator === '<=') {
            // Inclusive upper bound: <=17.0.0 means we can use exactly that version
            maxMajor = Math.max(maxMajor, comparator.semver.major);
            // Store the exact version for <= comparator
            if (highestMajor === null || comparator.semver.major > highestMajor) {
              exactVersion = comparator.semver.version;
            }
          }
          else if (comparator.operator === '>=' || comparator.operator === '>') {
            // For lower bounds use the major version
            maxMajor = Math.max(maxMajor, comparator.semver.major);
          }
        }
      }

      // Track the highest major version across all ranges
      highestMajor = Math.max(highestMajor ?? 0, maxMajor);
    }

    if (highestMajor === null) {
      throw new Error(`Unable to resolve the highest Node version for range '${versionRange}'.`);
    }

    // Return exact version if we have a <= comparator, otherwise use .x
    return exactVersion || `${highestMajor}.x`;
  }
};
