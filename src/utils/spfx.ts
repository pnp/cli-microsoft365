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
    const ranges = versionRange.split('||').map(r => r.trim());

    const versions = ranges.map(range => {
      if (range.includes('<')) {
        const upperBound = range.split('<')[1].trim().split(' ')[0];
        const parts = upperBound.split('.');
        return `${parts[0]}.${parts[1]}`;
      }

      const cleaned = range.replace(/[\^>=<~]/g, '');
      const parts = cleaned.split('.');

      if (parts.length >= 2) {
        return `${parts[0]}.${parts[1]}`;
      }

      return `${parts[0]}.0`;
    });

    const sorted = versions.sort((a, b) => {
      const [aMajor, aMinor] = a.split('.').map(Number);
      const [bMajor, bMinor] = b.split('.').map(Number);

      if (aMajor !== bMajor) {
        return bMajor - aMajor;
      }
      return bMinor - aMinor;
    });

    return `${sorted[0]}.x`;
  }
};