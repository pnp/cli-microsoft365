import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import * as fs from 'fs';
import * as path from 'path';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import { Group } from './Group';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description: string;
  mailNickname: string;
  owners?: string;
  members?: string;
  isPrivate?: string;
  logoPath?: string;
}

class AadO365GroupAddCommand extends GraphCommand {
  private static numRepeat: number = 15;

  public get name(): string {
    return commands.O365GROUP_ADD;
  }

  public get description(): string {
    return 'Creates Microsoft 365 Group';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let group: Group;

    if (this.verbose) {
      cmd.log(`Creating Microsoft 365 Group...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      json: true,
      body: {
        description: args.options.description,
        displayName: args.options.displayName,
        groupTypes: [
          "Unified"
        ],
        mailEnabled: true,
        mailNickname: args.options.mailNickname,
        securityEnabled: false,
        visibility: args.options.isPrivate == 'true' ? 'Private' : 'Public'
      }
    };

    request
      .post<Group>(requestOptions)
      .then((res: Group): Promise<void> => {
        group = res;

        if (!args.options.logoPath) {
          if (this.debug) {
            cmd.log('logoPath not set. Skipping');
          }

          return Promise.resolve();
        }

        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          cmd.log(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/groups/${group.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          body: fs.readFileSync(fullPath)
        };

        return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
          this.setGroupLogo(requestOptions, AadO365GroupAddCommand.numRepeat, resolve, reject, cmd);
        });
      })
      .then((): Promise<{ value: { id: string; }[] }> => {
        if (!args.options.owners) {
          if (this.debug) {
            cmd.log('Owners not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const owners: string[] = args.options.owners.split(',').map(o => o.trim());

        if (this.verbose) {
          cmd.log('Retrieving user information to set group owners...');
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
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
          url: `${this.resource}/v1.0/groups/${group.id}/owners/$ref`,
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
        if (!args.options.members) {
          if (this.debug) {
            cmd.log('Members not set. Skipping');
          }

          return Promise.resolve(undefined as any);
        }

        const members: string[] = args.options.members.split(',').map(o => o.trim());

        if (this.verbose) {
          cmd.log('Retrieving user information to set group members...');
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/users?$filter=${members.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
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
          url: `${this.resource}/v1.0/groups/${group.id}/members/$ref`,
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
        cmd.log(group);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  private setGroupLogo(requestOptions: any, retryLeft: number, resolve: () => void, reject: (err: any) => void, cmd: CommandInstance): void {
    request
      .put(requestOptions)
      .then((res: any): void => {
        resolve();
      }, (err: any): void => {
        if (--retryLeft > 0) {
          setTimeout(() => {
            this.setGroupLogo(requestOptions, retryLeft, resolve, reject, cmd);
          }, 500 * (AadO365GroupAddCommand.numRepeat - retryLeft));
        }
        else {
          reject(err);
        }
      });
  }

  private getImageContentType(imagePath: string): string {
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --displayName <displayName>',
        description: 'Display name for the Microsoft 365 Group'
      },
      {
        option: '-d, --description <description>',
        description: 'Description for the Microsoft 365 Group'
      },
      {
        option: '-m, --mailNickname <mailNickname>',
        description: 'Name to use in the group e-mail (part before the @)'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of Microsoft 365 Group owners'
      },
      {
        option: '--members [members]',
        description: 'Comma-separated list of Microsoft 365 Group members'
      },
      {
        option: '--isPrivate [isPrivate]',
        description: 'Set to true if the Microsoft 365 Group should be private and to false if it should be public (default)'
      },
      {
        option: '-l, --logoPath [logoPath]',
        description: 'Local path to the image file to use as group logo'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.owners) {
        let owners: string[] = args.options.owners.split(',').map(o => o.trim());
        for (let i = 0; i < owners.length; i++) {
          if (owners[i].indexOf('@') < 0) {
            return `${owners[i]} is not a valid userPrincipalName`;
          }
        }
      }

      if (args.options.members) {
        let members: string[] = args.options.members.split(',').map(m => m.trim());
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
    };
  }
}

module.exports = new AadO365GroupAddCommand();
