import { TokenStorage } from './TokenStorage';
import * as os from 'os';
import * as path from 'path';
import * as fs from 'fs';

export interface Services {
  [key: string]: string;
}

export interface TokensFile {
  services?: Services;
}

export class FileTokenStorage implements TokenStorage {
  private filePath: string = path.join(os.homedir(), '.o365cli-tokens.json');

  public get(service: string): Promise<string> {
    return new Promise<string>((resolve: (token: string) => void, reject: (error: any) => void): void => {
      if (!fs.existsSync(this.filePath)) {
        reject('File not found');
        return;
      }

      const contents: string = fs.readFileSync(this.filePath, 'utf8');
      try {
        const tokensFile: TokensFile = JSON.parse(contents);
        if (tokensFile.services &&
          tokensFile.services[service]) {
          resolve(tokensFile.services[service]);
        }
        else {
          reject(`Token for service ${service} not found`);
        }
      }
      catch (e) {
        reject(e);
      }
    });
  };

  public set(service: string, token: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      let tokensFile: TokensFile = {};

      if (fs.existsSync(this.filePath)) {
        const contents: string = fs.readFileSync(this.filePath, 'utf8');
        try {
          tokensFile = JSON.parse(contents);
        }
        catch (e) {
        }
      }

      if (!tokensFile.services) {
        tokensFile.services = {};
      }

      tokensFile.services[service] = token;

      fs.writeFile(this.filePath, JSON.stringify(tokensFile), 'utf8', (err: NodeJS.ErrnoException | null): void => {
        if (err) {
          reject(err.message);
        }
        else {
          resolve();
        }
      });
    });
  };

  public remove(service: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (!fs.existsSync(this.filePath)) {
        resolve();
        return;
      }

      let tokensFile: TokensFile = {};
      const contents: string = fs.readFileSync(this.filePath, 'utf8');
      try {
        tokensFile = JSON.parse(contents);
      }
      catch (e) {
        resolve();
        return;
      }

      if (!tokensFile.services) {
        resolve();
        return;
      }

      delete tokensFile.services[service];

      fs.writeFile(this.filePath, JSON.stringify(tokensFile), 'utf8', (err: NodeJS.ErrnoException | null): void => {
        if (err) {
          reject(err.message);
        }
        else {
          resolve();
        }
      });
    });
  };
}