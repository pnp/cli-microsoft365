import { Group, User } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { setTimeout } from 'timers/promises';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description: string;
  mailNickname: string;
  owners?: string;
  members?: string;
  isPrivate?: boolean;
  logoPath?: string;
  allowMembersToPost?: boolean;
  hideGroupInOutlook?: boolean;
  subscribeNewGroupMembers?: boolean;
  welcomeEmailDisabled?: boolean;
}

class AadO365GroupAddCommand extends GraphCommand {
  private static numRepeat: number = 15;
  private pollingInterval: number = 500;

  public get name(): string {
    return commands.O365GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft 365 Group';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        owners: typeof args.options.owners !== 'undefined',
        members: typeof args.options.members !== 'undefined',
        logoPath: typeof args.options.logoPath !== 'undefined',
        isPrivate: typeof args.options.isPrivate !== 'undefined',
        allowMembersToPost: args.options.allowMembersToPost,
        hideGroupInOutlook: args.options.hideGroupInOutlook,
        subscribeNewGroupMembers: args.options.subscribeNewGroupMembers,
        welcomeEmailDisabled: args.options.welcomeEmailDisabled
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '--isPrivate'
      },
      {
        option: '--allowMembersToPost [allowMembersToPost]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hideGroupInOutlook [hideGroupInOutlook]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--subscribeNewGroupMembers [subscribeNewGroupMembers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--welcomeEmailDisabled [welcomeEmailDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '-l, --logoPath [logoPath]'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowMembersToPost', 'hideGroupInOutlook', 'subscribeNewGroupMembers', 'welcomeEmailDisabled');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

    const requestOptions: CliRequestOptions = {
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
        visibility: args.options.isPrivate ? 'Private' : 'Public'
      }
    };

    try {
      ownerIds = await this.getUserIds(logger, args.options.owners);
      memberIds = await this.getUserIds(logger, args.options.members);
      group = await request.post<Group>(requestOptions);

      if (!args.options.logoPath) {
        if (this.debug) {
          logger.logToStderr('logoPath not set. Skipping');
        }
      }
      else {
        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          logger.logToStderr(`Setting group logo ${fullPath}...`);
        }

        const requestOptionsPhoto: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        await this.setGroupLogo(requestOptionsPhoto, AadO365GroupAddCommand.numRepeat, logger);
      }

      if (ownerIds.length !== 0) {
        await Promise.all(ownerIds.map(ownerId => request.post<void>({
          url: `${this.resource}/v1.0/groups/${group.id}/owners/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${ownerId}`
          }
        })));
      }

      if (memberIds.length !== 0) {
        await Promise.all(memberIds.map(memberId => request.post<void>({
          url: `${this.resource}/v1.0/groups/${group.id}/members/$ref`,
          headers: {
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${memberId}`
          }
        })));
      }

      logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserIds(logger: Logger, users: string | undefined): Promise<string[]> {
    if (!users) {
      if (this.debug) {
        logger.logToStderr('No users to validate, skipping.');
      }
      return [];
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving user information.');
    }

    const userArr: string[] = users.split(',').map(o => o.trim());
    let promises: Promise<{ value: User[] }>[] = [];
    let userIds: string[] = [];

    promises = userArr.map(user => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    });

    const usersRes = await Promise.all(promises);
    let userUpns: string[] = [];

    userUpns = usersRes.map(res => res.value[0]?.userPrincipalName as string);
    userIds = usersRes.map(res => res.value[0]?.id as string);

    // Find the members where no graph response was found
    const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));

    if (invalidUsers && invalidUsers.length > 0) {
      throw `Cannot proceed with group creation. The following users provided are invalid : ${invalidUsers.join(',')}`;
    }
    return userIds;
  }

  private async setGroupLogo(requestOptions: any, retryLeft: number, logger: Logger): Promise<void> {
    try {
      await request.put(requestOptions);
    }
    catch (err: any) {
      if (--retryLeft > 0) {
        await setTimeout(this.pollingInterval * (AadO365GroupAddCommand.numRepeat - retryLeft));
        await this.setGroupLogo(requestOptions, retryLeft, logger);
      }
      else {
        throw err;
      }
    }
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
}

module.exports = new AadO365GroupAddCommand();