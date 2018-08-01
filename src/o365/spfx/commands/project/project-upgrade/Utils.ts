import * as fs from 'fs';
import * as path from 'path';
import { Project } from './model';

export class Utils {
  public static removeSingleLineComments(s: string): string {
    const commentEval: RegExp = new RegExp(/\/\*[\s\S]*?\*\/|([^:]|^)\/\/.*$/gm);
    return s.replace(commentEval, '');
  }

  public static getAllFiles(dir: string): string | string[] {
    return fs.statSync(dir).isDirectory()
      ? Array.prototype.concat(...fs.readdirSync(dir).map(f => Utils.getAllFiles(path.join(dir, f))))
      : dir;
  }

  public static isReactProject(project: Project): boolean {
    return typeof project.yoRcJson !== 'undefined' &&
      typeof project.yoRcJson['@microsoft/generator-sharepoint'] !== 'undefined' &&
      project.yoRcJson["@microsoft/generator-sharepoint"].framework === 'react';
  }
}