import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import * as fs from 'fs';
import * as path from 'path';
const vorpal: Vorpal = require('../../../vorpal-init');

export interface Options extends GlobalOptions {
  id: string;
  displayName?: string;
  description?: string;
  owners?: string;
  members?: string;
  mailNickname?: string;
  classification?: string;
  isPrivate?: string;
  logoPath?: string;
}
export class GroupUpdateService {

  private static numRepeat: number = 15;

  public static UpdateGroup(cmd: CommandInstance, 
    resource: string, 
    options: Options,
    verbose: boolean,
    debug: boolean, 
    successCallback: () => void, 
    errorCallback: (rawRes: any, cmd: CommandInstance, cb: () => void) => void): void {
    ((): Promise<void> => {
      if (!options.displayName &&
        !options.description &&
        typeof options.isPrivate === 'undefined') {
        return Promise.resolve();
      }

      if (verbose) {
        cmd.log(`Updating Office 365 Group ${options.id}...`);
      }

      const update: any = {};
      if (options.displayName) {
        update.displayName = options.displayName;
      }
      if (options.description) {
        update.description = options.description;
      }
      if (typeof options.isPrivate !== 'undefined') {
        update.visibility = options.isPrivate == 'true' ? 'Private' : 'Public'
      }

      const requestOptions: any = {
        url: `${resource}/v1.0/groups/${options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        json: true,
        body: update
      };

      return request.patch(requestOptions);
    })()
      .then((): Promise<void> => {
        if (!options.logoPath) {
          if (debug) {
            cmd.log('logoPath not set. Skipping');
          }

          return Promise.resolve();
        }

        const fullPath: string = path.resolve(options.logoPath);
        if (verbose) {
          cmd.log(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: any = {
          url: `${resource}/v1.0/groups/${options.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          body: fs.readFileSync(fullPath)
        };

        return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
          this.setGroupLogo(requestOptions, GroupUpdateService.numRepeat, debug, resolve, reject, cmd);
        });
      })
      .then((): Promise<{ value: { id: string; }[] }> => {
        if (!options.owners) {
          if (debug) {
            cmd.log('Owners not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const owners: string[] = options.owners.split(',').map(o => o.trim());

        if (verbose) {
          cmd.log('Retrieving user information to set group owners...');
        }

        const requestOptions: any = {
          url: `${resource}/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res?: { value: { id: string; }[] }): Promise<any> => {
        if (!res) {
          return Promise.resolve();
        }

        return Promise.all(res.value.map(u => request.post({
          url: `${resource}/v1.0/groups/${options.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          json: true,
          body: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      })
      .then((): Promise<{ value: { id: string; }[] }> => {
        if (!options.members) {
          if (debug) {
            cmd.log('Members not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const members: string[] = options.members.split(',').map(o => o.trim());

        if (verbose) {
          cmd.log('Retrieving user information to set group members...');
        }

        const requestOptions: any = {
          url: `${resource}/v1.0/users?$filter=${members.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res?: { value: { id: string; }[] }): Promise<any> => {
        if (!res) {
          return Promise.resolve();
        }

        return Promise.all(res.value.map(u => request.post({
          url: `${resource}/v1.0/groups/${options.id}/members/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          json: true,
          body: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      })
      .then((): void => {
        if (verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        successCallback();
      }, (rawRes: any): void => errorCallback(rawRes, cmd, successCallback));
  }
  private static setGroupLogo(requestOptions: any, 
    retryLeft: number, 
    debug: boolean,
    resolve: () => void, 
    reject: (err: any) => void, 
    cmd: CommandInstance): void {
    request
      .put(requestOptions)
      .then((res: any): void => {
        if (debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        resolve();
      }, (err: any): void => {
        if (--retryLeft > 0) {
          setTimeout(() => {
            this.setGroupLogo(requestOptions, retryLeft, debug, resolve, reject, cmd);
          }, 500 * (GroupUpdateService.numRepeat - retryLeft));
        }
        else {
          reject(err);
        }
      });
  }
  private static getImageContentType(imagePath: string): string {
    let extension: string = imagePath.substr(imagePath.lastIndexOf('.')).toLowerCase();

    switch (extension) {
      case '.png':
        return 'image/png';
      case '.gif':
        return 'image/gif';
      default:
        return 'image/jpeg';
    }
  }
}