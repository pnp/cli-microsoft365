import { minVersion, SemVer, validRange } from 'semver';
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

    let highest: { version: SemVer; source: string } | null = null;

    for (const range of ranges) {
      const normalized = validRange(range);
      if (!normalized) {
        continue;
      }

      const minimum = minVersion(normalized);
      if (!minimum) {
        continue;
      }

      if (!highest || minimum.compare(highest.version) > 0) {
        highest = {
          version: minimum,
          source: range
        };
      }
    }

    if (!highest) {
      throw new Error(`Unable to resolve the highest Node version for range '${versionRange}'.`);
    }

    const source = highest.source.trim();
    const compactSource = source.replace(/\s+/g, '');
    const isCaretOrTilde = compactSource.startsWith('^') || compactSource.startsWith('~');
    const isSimpleVersion = /^[0-9]+(\.[0-9]+){0,2}$/.test(compactSource);

    if (isCaretOrTilde || isSimpleVersion) {
      const numeric = isCaretOrTilde ? compactSource.substring(1) : compactSource;
      const parts = numeric.split('.').filter(part => part.length > 0);
      const { major } = highest.version;

      if (parts.length >= 2) {
        return `${major}.${parts[1]}.x`;
      }

      return `${major}.x`;
    }

    return highest.version.version;
  }
};
