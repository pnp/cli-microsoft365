import fs from 'fs';
import os from 'os';
import path from 'path';
import { TokenStorage } from './TokenStorage.js';

export class FileTokenStorage implements TokenStorage {
  public static msalCacheFilePath(): string {
    return path.join(os.homedir(), '.cli-m365-msal.json');
  }

  public static connectionInfoFilePath(): string {
    return path.join(os.homedir(), '.cli-m365-tokens.json');
  }

  constructor(private filePath: string) {
  }

  public async get(): Promise<string> {
    if (!fs.existsSync(this.filePath)) {
      throw 'File not found';
    }
    const contents: string = fs.readFileSync(this.filePath, 'utf8');
    return contents;
  }

  public async set(connectionInfo: string): Promise<void> {
    return fs.writeFile(this.filePath, connectionInfo, 'utf8', (err: NodeJS.ErrnoException | null): void => {
      if (err) {
        throw err.message;
      }
      else {
        return;
      }
    });
  }

  public async remove(): Promise<void> {
    if (!fs.existsSync(this.filePath)) {
      return;
    }

    return fs.writeFile(this.filePath, '', 'utf8', (err: NodeJS.ErrnoException | null): void => {
      if (err) {
        throw err.message;
      }
      else {
        return;
      }
    });
  }
}