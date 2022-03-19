import { Group, User } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
  allowMembersToPost?: boolean;
  hideGroupInOutlook?: boolean;
  subscribeNewGroupMembers?: boolean;
  welcomeEmailDisabled?: boolean;
}

class AadO365GroupAddCommand extends GraphCommand {
  private static numRepeat: number = 15;

  public get name(): string {
    return commands.O365GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.owners = typeof args.options.owners !== 'undefined';
    telemetryProps.members = typeof args.options.members !== 'undefined';
    telemetryProps.logoPath = typeof args.options.logoPath !== 'undefined';
    telemetryProps.isPrivate = typeof args.options.isPrivate !== 'undefined';
    telemetryProps.allowMembersToPost = (!(!args.options.allowMembersToPost)).toString();
    telemetryProps.hideGroupInOutlook = (!(!args.options.hideGroupInOutlook)).toString();
    telemetryProps.subscribeNewGroupMembers = (!(!args.options.subscribeNewGroupMembers)).toString();
    telemetryProps.welcomeEmailDisabled = (!(!args.options.welcomeEmailDisabled)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let group: Group;
    let ownerIds: string[] = [];
    let memberIds: string[] = [];
    const resourceBehaviorOptionsCollection: string[] = [];

    if (this.verbose) {
      logger.logToStderr(`Creating Microsoft 365 Group...`);
    }

    if (args.options.allowMembersToPost) {
      resourceBehaviorOptionsCollection.push("allowMembersToPost");
    }

    if (args.options.hideGroupInOutlook) {
      resourceBehaviorOptionsCollection.push("hideGroupInOutlook");
    }

    if (args.options.subscribeNewGroupMembers) {
      resourceBehaviorOptionsCollection.push("subscribeNewGroupMembers");
    }

    if (args.options.welcomeEmailDisabled) {
      resourceBehaviorOptionsCollection.push("welcomeEmailDisabled");
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
        resourceBehaviorOptions: resourceBehaviorOptionsCollection,
        securityEnabled: false,
        visibility: args.options.isPrivate === 'true' ? 'Private' : 'Public'
      }
    };

    this
      .getUserIds(logger, args.options.owners)
      .then((ownerIdsRes: string[]): Promise<string[]> => {
        ownerIds = ownerIdsRes;
        return this.getUserIds(logger, args.options.members);
      })
      .then((memberIdsRes: string[]): Promise<Group> => {
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
      .then((): Promise<void[]> => {
        if (ownerIds.length === 0) {
          return Promise.resolve([]);
        }

        return Promise.all(ownerIds.map(ownerId => request.post<void>({
          url: `${this.resource}/v1.0/groups/${group.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${ownerId}`
          }
        })));
      })
      .then((): Promise<void[]> => {
        if (memberIds.length === 0) {
          return Promise.resolve([]);
        }

        return Promise.all(memberIds.map(memberId => request.post<void>({
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

  private getUserIds(logger: Logger, users: string | undefined): Promise<string[]> {
    if (!users) {
      if (this.debug) {
        logger.logToStderr('No users to validate, skipping.');
      }
      return Promise.resolve([]);
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving user information.');
    }

    const userArr: string[] = users.split(',').map(o => o.trim());
    let promises: Promise<{ value: User[] }>[] = [];
    let userIds: string[] = [];

    promises = userArr.map(user => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    return Promise.all(promises).then((usersRes: { value: User[] }[]): Promise<string[]> => {
      let userUpns: string[] = [];

      userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
      userIds = usersRes.map(res => res.value[0]?.id as string);

      // Find the members where no graph response was found
      const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

      if (invalidUsers && invalidUsers.length > 0) {
        return Promise.reject(`Cannot proceed with group creation. The following users provided are invalid : ${invalidUsers.join(',')}`);
      }
      return Promise.resolve(userIds);
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
        option: '--allowMembersToPost [allowMembersToPost]'
      },
      {
        option: '--hideGroupInOutlook [hideGroupInOutlook]'
      },
      {
        option: '--subscribeNewGroupMembers [subscribeNewGroupMembers]'
      },
      {
        option: '--welcomeEmailDisabled [welcomeEmailDisabled]'
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