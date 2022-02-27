import { Group } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id: string;
  displayName?: string;
  description?: string;
  owners?: string;
  members?: string;
  isPrivate?: string;
  logoPath?: string;
}

class AadO365GroupSetCommand extends GraphCommand {
  private static numRepeat: number = 15;

  public get name(): string {
    return commands.O365GROUP_SET;
  }

  public get description(): string {
    return 'Updates Microsoft 365 Group properties';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    ((): Promise<void> => {
      if (!args.options.displayName &&
        !args.options.description &&
        typeof args.options.isPrivate === 'undefined') {
        return Promise.resolve();
      }

      if (this.verbose) {
        logger.logToStderr(`Updating Microsoft 365 Group ${args.options.id}...`);
      }

      const update: Group = {};
      if (args.options.displayName) {
        update.displayName = args.options.displayName;
      }
      if (args.options.description) {
        update.description = args.options.description;
      }
      if (typeof args.options.isPrivate !== 'undefined') {
        update.visibility = args.options.isPrivate === 'true' ? 'Private' : 'Public';
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/groups/${args.options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: update
      };

      return request.patch(requestOptions);
    })()
      .then((): Promise<void> => {
        if (!args.options.logoPath) {
          if (this.debug) {
            logger.logToStderr('logoPath not set. Skipping');
          }

          return Promise.resolve();
        }

        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          logger.logToStderr(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/groups/${args.options.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
          this.setGroupLogo(requestOptions, AadO365GroupSetCommand.numRepeat, resolve, reject, logger);
        });
      })
      .then((): Promise<{ value: { id: string; }[] }> => {
        if (!args.options.owners) {
          if (this.debug) {
            logger.logToStderr('Owners not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const owners: string[] = args.options.owners.split(',').map(o => o.trim());

        if (this.verbose) {
          logger.logToStderr('Retrieving user information to set group owners...');
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res?: { value: { id: string; }[] }): Promise<any> => {
        if (!res) {
          return Promise.resolve();
        }

        return Promise.all(res.value.map(u => request.post({
          url: `${this.resource}/v1.0/groups/${args.options.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      })
      .then((): Promise<{ value: { id: string; }[] }> => {
        if (!args.options.members) {
          if (this.debug) {
            logger.logToStderr('Members not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const members: string[] = args.options.members.split(',').map(o => o.trim());

        if (this.verbose) {
          logger.logToStderr('Retrieving user information to set group members...');
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/users?$filter=${members.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res?: { value: { id: string; }[] }): Promise<any> => {
        if (!res) {
          return Promise.resolve();
        }

        return Promise.all(res.value.map(u => request.post({
          url: `${this.resource}/v1.0/groups/${args.options.id}/members/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${u.id}`
          }
        })));
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private setGroupLogo(requestOptions: any, retryLeft: number, resolve: () => void, reject: (err: any) => void, logger: Logger): void {
    request
      .put(requestOptions)
      .then((res: any): void => {
        if (this.debug) {
          logger.logToStderr('Response:');
          logger.logToStderr(res);
          logger.logToStderr('');
        }

        resolve();
      }, (err: any): void => {
        if (--retryLeft > 0) {
          setTimeout(() => {
            this.setGroupLogo(requestOptions, retryLeft, resolve, reject, logger);
          }, 500 * (AadO365GroupSetCommand.numRepeat - retryLeft));
        }
        else {
          reject(err);
        }
      });
  }

  private getImageContentType(imagePath: string): string {
    const extension: string = imagePath.substr(imagePath.lastIndexOf('.')).toLowerCase();

    switch (extension) {
      case '.png':
        return 'image/png';
      case '.gif':
        return 'image/gif';
      default:
        return 'image/jpeg';
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--owners [owners]'
      },
      {
        option: '--members [members]'
      },
      {
        option: '--isPrivate [isPrivate]'
      },
      {
        option: '-l, --logoPath [logoPath]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.displayName &&
      !args.options.description &&
      !args.options.members &&
      !args.options.owners &&
      typeof args.options.isPrivate === 'undefined' &&
      !args.options.logoPath) {
      return 'Specify at least one property to update';
    }

    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.owners) {
      const owners: string[] = args.options.owners.split(',').map(o => o.trim());
      for (let i = 0; i < owners.length; i++) {
        if (owners[i].indexOf('@') < 0) {
          return `${owners[i]} is not a valid userPrincipalName`;
        }
      }
    }

    if (args.options.members) {
      const members: string[] = args.options.members.split(',').map(m => m.trim());
      for (let i = 0; i < members.length; i++) {
        if (members[i].indexOf('@') < 0) {
          return `${members[i]} is not a valid userPrincipalName`;
        }
      }
    }

    if (typeof args.options.isPrivate !== 'undefined' &&
      args.options.isPrivate !== 'true' &&
      args.options.isPrivate !== 'false') {
      return `${args.options.isPrivate} is not a valid boolean value`;
    }

    if (args.options.logoPath) {
      const fullPath: string = path.resolve(args.options.logoPath);

      if (!fs.existsSync(fullPath)) {
        return `File '${fullPath}' not found`;
      }

      if (fs.lstatSync(fullPath).isDirectory()) {
        return `Path '${fullPath}' points to a directory`;
      }
    }

    return true;
  }
}

module.exports = new AadO365GroupSetCommand();
