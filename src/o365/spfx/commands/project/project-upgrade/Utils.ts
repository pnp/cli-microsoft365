import * as fs from 'fs';
import * as path from 'path';

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
}