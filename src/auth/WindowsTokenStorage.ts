import { TokenStorage } from './TokenStorage';
import * as childProcess from 'child_process';
import * as os from 'os';
import * as path from 'path';
import { Buffer } from 'buffer';

interface StorageEntry {
  name: string;
  value: string;
}

export class WindowsTokenStorage implements TokenStorage {
  private credsExePath: string = path.join(__dirname, '../../bin/windows/creds.exe');
  private prefix: string = 'Office365Cli:target=';
  private prefixShort: string = 'Office365Cli';
  private MAX_CREDENTIAL_BYTES: number = 2048;
  private propertyRegex: RegExp = /^([^:]+):\s(.*)$/;

  public get(): Promise<string> {
    return new Promise<string>((resolve: (connectionInfo: string) => void, reject: (error: any) => void): void => {
      const args: string[] = [
        '-s',
        '-g',
        '-t', `${this.prefix}${this.prefixShort}*`
      ];

      childProcess.execFile(this.credsExePath, args, (err: Error | null, stdout: string, stderr: string): void => {
        if (err) {
          reject(err.message);
          return;
        }

        const lines: string[] = stdout.split(os.EOL);
        const creds: StorageEntry[] = [];
        let cred: StorageEntry = { name: '', value: '' };
        lines.forEach(l => {
          // empty line is a separator, so reset object
          if (l === '') {
            cred = { name: '', value: '' };
          }

          const m: RegExpExecArray | null = this.propertyRegex.exec(l);
          if (!m) {
            return;
          }

          switch (m[1]) {
            case 'Target Name':
              cred.name = m[2];
              break;
            case 'Credential':
              cred.value = m[2];
              break;
          }

          if (cred.name.length > 0 && cred.value.length > 0) {
            creds.push(cred);
          }
        });

        if (creds.length === 0) {
          reject('Credential not found');
          return;
        }

        let rawPassword: string = '';
        if (creds.length === 1 && !this.isPartialEntry(creds[0].name)) {
          rawPassword = creds[0].value;
        }
        else {
          const chunks: string[] = [];
          let numChunks: number = 0;
          creds.forEach(c => {
            const chunkInfo: number[] = this.getChunkInfo(c.name);
            if (chunkInfo.length !== 2) {
              return;
            }

            if (chunkInfo[0] === 1) {
              numChunks = chunkInfo[1];
            }

            chunks[chunkInfo[0]] = c.value;
          });
          if (chunks.length - 1 !== numChunks) {
            reject(`Couldn't load all credential chunks. Expected ${numChunks}, found ${chunks.length - 1}`);
            return;
          }

          for (let i: number = 1; i < chunks.length; i++) {
            if (!chunks[i]) {
              reject(`Missing chunk ${i}/${numChunks}`);
              return;
            }
          }

          rawPassword = chunks.join('');
        }

        resolve(new Buffer(rawPassword, 'hex').toString('utf8'));
      });
    });
  };

  public set(connectionInfo: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      // because the new token might be longer/shorter than the previous one
      // we first need to clear the previous entry to avoid issues
      this
        .remove()
        .then((): void => {
          const entries: StorageEntry[] = [];

          if (connectionInfo.length <= this.MAX_CREDENTIAL_BYTES) {
            entries.push({ name: this.prefix + this.prefixShort, value: connectionInfo });
          }
          else {
            const numBytes: number = connectionInfo.length;
            let numBlocks: number = Math.ceil(numBytes / this.MAX_CREDENTIAL_BYTES);

            for (let i: number = 0; i < numBlocks; i++) {
              entries.push({
                name: `${this.prefix}${this.prefixShort}--${i + 1}-${numBlocks}`,
                value: connectionInfo.substr(i * this.MAX_CREDENTIAL_BYTES, this.MAX_CREDENTIAL_BYTES)
              });
            }
          }

          let i: number = 0;
          entries.forEach(e => {
            const args: string[] = [
              '-a',
              '-t', e.name,
              '-p', new Buffer(e.value as string, 'utf8').toString('hex')
            ];

            childProcess.execFile(this.credsExePath, args, (err: Error | null, stdout: string, stderr: string): void => {
              if (err) {
                reject('Could not add password to credential store: ' + err.message);
              }
              else {
                ++i;
                if (i === entries.length) {
                  resolve();
                }
              }
            });
          });
        }, (error: any): void => {
          reject(error);
        });
    });
  };

  public remove(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const args: string[] = [
        '-d',
        '-g',
        '-t', `${this.prefix}${this.prefixShort}*`
      ];

      childProcess.execFile(this.credsExePath, args, (err: Error | null, stdout: string, stderr: string): void => {
        if (err) {
          reject('Could not remove password from credential store: ' + err.message);
        }
        else {
          resolve();
        }
      });
    });
  };

  private isPartialEntry(entryName: string): boolean {
    return /--\d+-\d+$/.test(entryName);
  }

  private getChunkInfo(entryName: string): number[] {
    const m: RegExpExecArray | null = /--(\d+)-(\d+)$/.exec(entryName);
    if (m) {
      return [parseInt(m[1]), parseInt(m[2])];
    }
    else {
      return [];
    }
  }
}