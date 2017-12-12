import { TokenStorage } from './TokenStorage';
import * as childProcess from 'child_process';

export class KeychainTokenStorage implements TokenStorage {
  private securityPath: string = '/usr/bin/security';
  private description: string = 'Office 365 CLI';

  public get(service: string): Promise<string> {
    return new Promise<string>((resolve: (token: string) => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'find-generic-password',
        '-a', service,
        '-s', service,
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

  public set(service: string, token: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'add-generic-password',
        '-a', service,
        '-s', service,
        '-D', this.description,
        '-w', token,
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

  public remove(service: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const args: string[] = [
        'delete-generic-password',
        '-a', service,
        '-s', service,
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