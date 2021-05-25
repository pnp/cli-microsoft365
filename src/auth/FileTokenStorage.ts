import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { TokenStorage } from './TokenStorage';

export class FileTokenStorage implements TokenStorage {
  public static msalCacheFilePath(): string {
    return path.join(os.homedir(), '.cli-m365-msal.json');
  }
  
  public static connectionInfoFilePath(): string {
    return path.join(os.homedir(), '.cli-m365-tokens.json');
  }

  constructor(private filePath: string) {
  }

  public get(): Promise<string> {
    return new Promise<string>((resolve: (connectionInfo: string) => void, reject: (error: any) => void): void => {
      if (!fs.existsSync(this.filePath)) {
        reject('File not found');
        return;
      }

      const contents: string = fs.readFileSync(this.filePath, 'utf8');
      resolve(contents);
    });
  }

  public set(connectionInfo: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      fs.writeFile(this.filePath, connectionInfo, 'utf8', (err: NodeJS.ErrnoException | null): void => {
        if (err) {
          reject(err.message);
        }
        else {
          resolve();
        }
      });
    });
  }

  public remove(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (!fs.existsSync(this.filePath)) {
        resolve();
        return;
      }

      fs.writeFile(this.filePath, '', 'utf8', (err: NodeJS.ErrnoException | null): void => {
        if (err) {
          reject(err.message);
        }
        else {
          resolve();
        }
      });
    });
  }
}