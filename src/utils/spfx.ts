import { Project } from '../m365/spfx/commands/project/project-model';

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
  }
};