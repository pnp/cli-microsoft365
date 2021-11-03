import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Group } from './Group';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let group: Group;
    let ownerIds: [] = [];
    let memberIds: [] = [];

    if (this.verbose) {
      logger.logToStderr(`Creating Microsoft 365 Group...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.displayName,
        groupTypes: [
          "Unified"
        ],
        mailEnabled: true,
        mailNickname: args.options.mailNickname,
        securityEnabled: false,
        visibility: args.options.isPrivate === 'true' ? 'Private' : 'Public'
      }
    };

    this.validateOwners(logger, args)
      .then((ownerIdsRes: []): Promise<any> => {
        ownerIds = ownerIdsRes;
        return this.validateMembers(logger, args);
      })
      .then((memberIdsRes: []): Promise<Group> => {
        memberIds = memberIdsRes;
        return request.post<Group>(requestOptions);
      })
      .then((res: Group): Promise<void> => {
        group = res;

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
          url: `${this.resource}/v1.0/groups/${group.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
          this.setGroupLogo(requestOptions, AadO365GroupAddCommand.numRepeat, resolve, reject, logger);
        });
      })
      .then((): Promise<any> => {
        if (!ownerIds || ownerIds.length === 0) {
          return Promise.resolve();
        }

        return Promise.all(ownerIds.map(ownrId => request.post({
          url: `${this.resource}/v1.0/groups/${group.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${ownrId}`
          }
        })));
      })
      .then((): Promise<any> => {
        if (!memberIds || memberIds.length === 0) {
          return Promise.resolve();
        }

        return Promise.all(memberIds.map(memberId => request.post({
          url: `${this.resource}/v1.0/groups/${group.id}/members/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${memberId}`
          }
        })));
      })
      .then((): void => {
        logger.log(group);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }


  private validateOwners(logger: Logger, args: CommandArgs): Promise<any> {
    if (!args.options.owners) {
      if (this.debug) {
        logger.logToStderr('Owners not set. Skipping');
      }
      return Promise.resolve(undefined as any);
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving user information to set group owners...');
    }

    const owners: string[] = args.options.owners.split(',').map(o => o.trim());
    const promises: any[] = [];
    let ownerIds: any[] = [];

    owners.forEach(owner => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=${`userPrincipalName eq '${owner}'`}&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      promises.push(request.get(requestOptions));
    });

    return Promise.all(promises).then((ownersRes: { value: { id: string; userPrincipalName: string }[] }[]): Promise<any> => {
      let ownerUpns: any[] = [];

      ownerUpns = ownersRes.map(res => res.value[0]?.userPrincipalName);
      ownerIds = ownersRes.map(res => res.value[0]?.id);

      // Find the owners where no graph response was found
      const invalidUsers = owners.filter(ownr => !ownerUpns.some((upn) => upn === ownr));

      if (invalidUsers && invalidUsers.length > 0) {
        return Promise.reject(`Cannot proceed with group creation. The following Owners provided are invalid : ${invalidUsers.join(',')}`);
      }
      return Promise.resolve(ownerIds);
    });
  }

  private validateMembers(logger: Logger, args: CommandArgs): Promise<any> {
    if (!args.options.members) {
      if (this.debug) {
        logger.logToStderr('members not set. Skipping');
      }
      return Promise.resolve(undefined as any);
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving user information to set group members...');
    }

    const members: string[] = args.options.members.split(',').map(o => o.trim());
    const promises: any[] = [];
    let memberIds: any[] = [];

    members.forEach(member => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=${`userPrincipalName eq '${member}'`}&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      promises.push(request.get(requestOptions));
    });

    return Promise.all(promises).then((membersRes: { value: { id: string; userPrincipalName: string }[] }[]): Promise<any> => {
      let memberUpns: any[] = [];

      memberUpns = membersRes.map(res => res.value[0]?.userPrincipalName);
      memberIds = membersRes.map(res => res.value[0]?.id);

      // Find the members where no graph response was found
      const invalidUsers = members.filter(ownr => !memberUpns.some((upn) => upn === ownr));

      if (invalidUsers && invalidUsers.length > 0) {
        return Promise.reject(`Cannot proceed with group creation. The following Members provided are invalid : ${invalidUsers.join(',')}`);
      }
      return Promise.resolve(memberIds);
    });
  }

  private setGroupLogo(requestOptions: any, retryLeft: number, resolve: () => void, reject: (err: any) => void, logger: Logger): void {
    request
      .put(requestOptions)
      .then((): void => resolve(),
        (err: any): void => {
          if (--retryLeft > 0) {
            setTimeout(() => {
              this.setGroupLogo(requestOptions, retryLeft, resolve, reject, logger);
            }, 500 * (AadO365GroupAddCommand.numRepeat - retryLeft));
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
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-d, --description <description>'
      },
      {
        option: '-m, --mailNickname <mailNickname>'
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

module.exports = new AadO365GroupAddCommand();
