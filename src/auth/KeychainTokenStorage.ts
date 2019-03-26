import { TokenStorage } from './TokenStorage';
import * as childProcess from 'child_process';

export class KeychainTokenStorage implements TokenStorage {
  private securityPath: string = '/usr/bin/security';
  private description: string = 'Office 365 CLI';

  public get(): Promise<string> {
    return new Promise<string>((resolve: (connectionInfo: string) => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'find-generic-password',
        '-a', this.description,
        '-s', this.description,
        '-D', this.description,
        '-g'
      ];

      childProcess.execFile(this.securityPath, args, (err: Error | null, stdout: string, stderr: string): void => {
        if (err) {
          reject(err.message);
          return;
        }

        const match: RegExpExecArray | null = /^password: (?:0x[0-9A-F]+  )?"(.*)"$/m.exec(stderr);
        if (match) {
          const password: string = match[1].replace(/\\134/g, '\\');
          resolve(password);
          return;
        }

        reject('Password in invalid format');
      });
    });
  };

  public set(connectionInfo: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'add-generic-password',
        '-a', this.description,
        '-s', this.description,
        '-D', this.description,
        '-w', connectionInfo,
        '-U'
      ];

      childProcess.execFile(this.securityPath, args, (err: Error | null, stdout: string, stderr: string): void => {
        if (err) {
          reject('Could not add password to keychain: ' + err.message);
        }
        else {
          resolve();
        }
      });
    });
  };

  public remove(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'delete-generic-password',
        '-a', this.description,
        '-s', this.description,
        '-D', this.description
      ];

      childProcess.execFile(this.securityPath, args, (err: Error | null, stdout: string, stderr: string): void => {
        if (err) {
          reject('Could not remove account from keychain: ' + err.message);
        }
        else {
          resolve();
        }
      });
    });
  };
}