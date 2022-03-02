import * as fs from 'fs';
import * as path from 'path';

export const fsUtil = {
  readdirR(dir: string): string | string[] {
    return fs.statSync(dir).isDirectory()
      ? Array.prototype.concat(...fs.readdirSync(dir).map(f => fsUtil.readdirR(path.join(dir, f))))
      : dir;
  },

  // from: https://stackoverflow.com/a/22185855
  copyRecursiveSync(src: string, dest: string, replaceTokens?: (s: string) => string): void {
    const exists: boolean = fs.existsSync(src);
    const stats: false | fs.Stats = exists && fs.statSync(src);
    const isDirectory: boolean = exists && (stats as fs.Stats).isDirectory();

    if (replaceTokens) {
      dest = replaceTokens(dest);
    }

    if (isDirectory) {
      if (!fs.existsSync(dest)) {
        fs.mkdirSync(dest);
      }
      fs.readdirSync(src).forEach(function (childItemName) {
        fsUtil.copyRecursiveSync(path.join(src, childItemName),
          path.join(dest, childItemName), replaceTokens);
      });
    }
    else {
      fs.copyFileSync(src, dest);
    }
  },
  
  getSafeFileName(input: string): string {
    return input.replace(/'/g, "''");
  }
};