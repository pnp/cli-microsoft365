import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import config from '../config';

const cacheFolderPath = path.join(os.tmpdir(), config.configstoreName.replace('config', 'cache'));
const mkdirOptions = { mode: 0o0700, recursive: true };

export const cache = {
  cacheFolderPath: cacheFolderPath,
  getValue(key: string): string | undefined {
    this.clearExpired();

    try {
      const cacheFilePath = path.join(cacheFolderPath, key);
      if (!fs.existsSync(cacheFilePath)) {
        return undefined;
      }

      return fs.readFileSync(cacheFilePath, 'utf8');
    }
    catch {
      return undefined;
    }
  },

  setValue(key: string, value: string): void {
    this.clearExpired();

    try {
      fs.mkdirSync(cacheFolderPath, mkdirOptions);
      const cacheFilePath = path.join(cacheFolderPath, key);
      // we don't need to wait for the file to be written
      // eslint-disable-next-line @typescript-eslint/no-empty-function
      fs.writeFile(cacheFilePath, value, () => { });
    }
    catch { }
  },

  clearExpired(cb?: () => void): void {
    // we don't need to wait for this to complete
    // even if it stops meanwhile, it will be picked up next time
    fs.readdir(cacheFolderPath, (err, files) => {
      if (err) {
        if (cb) {
          cb();
        }
        return;
      }

      const numFiles: number = files.length;
      if (numFiles === 0) {
        if (cb) {
          cb();
        }
        return;
      }

      files.forEach((file, index) => {
        fs.stat(path.join(cacheFolderPath, file), (err, stats) => {
          if (err || stats.isDirectory()) {
            if (cb && index === numFiles - 1) {
              cb();
            }
            return;
          }

          // remove files that haven't been accessed in the last 24 hours
          if (stats.atime.getTime() < Date.now() - 24 * 60 * 60 * 1000) {
            // we don't need to wait for the file to be deleted
            // eslint-disable-next-line @typescript-eslint/no-empty-function
            fs.unlink(path.join(cacheFolderPath, file), () => {
              if (cb && index === numFiles - 1) {
                cb();
              }
            });
          }
          else {
            if (cb && index === numFiles - 1) {
              cb();
            }
          }
        });
      });
    });
  }
};