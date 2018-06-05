import * as fs from 'fs';
import * as path from 'path';
const stripJsonComments = require('strip-json-comments');

export class Utils {
  public static removeSingleLineComments(s: string): string {
    return stripJsonComments(s);
  }

  public static getAllFiles(dir: string): string | string[] {
    return fs.statSync(dir).isDirectory()
      ? Array.prototype.concat(...fs.readdirSync(dir).map(f => Utils.getAllFiles(path.join(dir, f))))
      : dir;
  }
}