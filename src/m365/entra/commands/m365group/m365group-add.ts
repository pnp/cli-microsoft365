import { Group, User } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description: string;
  mailNickname: string;
  owners?: string;
  members?: string;
  visibility?: string;
  logoPath?: string;
  allowMembersToPost?: boolean;
  hideGroupInOutlook?: boolean;
  subscribeNewGroupMembers?: boolean;
  welcomeEmailDisabled?: boolean;
}

class EntraM365GroupAddCommand extends GraphCommand {
  private static numRepeat: number = 15;
  private pollingInterval: number = 500;
  private allowedVisibilities: string[] = ['Private', 'Public', 'HiddenMembership'];

  public get name(): string {
    return commands.M365GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft 365 Group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_ADD];
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
        visibility: typeof args.options.visibility !== 'undefined',
        allowMembersToPost: !!args.options.allowMembersToPost,
        hideGroupInOutlook: !!args.options.hideGroupInOutlook,
        subscribeNewGroupMembers: !!args.options.subscribeNewGroupMembers,
        welcomeEmailDisabled: !!args.options.welcomeEmailDisabled
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
        option: '--visibility [visibility]',
        autocomplete: this.allowedVisibilities
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

        if (args.options.mailNickname.indexOf(' ') !== -1) {
          return 'The option mailNickname cannot contain spaces.';
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

        if (args.options.visibility && this.allowedVisibilities.map(x => x.toLowerCase()).indexOf(args.options.visibility.toLowerCase()) === -1) {
          return `${args.options.visibility} is not a valid visibility. Allowed values are ${this.allowedVisibilities.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.showDeprecationWarning(logger, aadCommands.M365GROUP_ADD, commands.M365GROUP_ADD);

    let group: Group;
    let ownerIds: string[] = [];
    let memberIds: string[] = [];
    const resourceBehaviorOptionsCollection: string[] = [];
    const resolvedVisibility = args.options.visibility || 'Public';

    if (this.verbose) {
      await logger.logToStderr('Creating Microsoft 365 Group...');
    }

    if (args.options.allowMembersToPost) {
      resourceBehaviorOptionsCollection.push('AllowOnlyMembersToPost');
    }

    if (args.options.hideGroupInOutlook) {
      resourceBehaviorOptionsCollection.push('HideGroupInOutlook');
    }

    if (args.options.subscribeNewGroupMembers) {
      resourceBehaviorOptionsCollection.push('SubscribeNewGroupMembers');
    }

    if (args.options.welcomeEmailDisabled) {
      resourceBehaviorOptionsCollection.push('WelcomeEmailDisabled');
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
        visibility: resolvedVisibility
      }
    };

    try {
      ownerIds = await this.getUserIds(logger, args.options.owners);
      memberIds = await this.getUserIds(logger, args.options.members);
      group = await request.post<Group>(requestOptions);

      if (!args.options.logoPath) {
        if (this.debug) {
          await logger.logToStderr('logoPath not set. Skipping');
        }
      }
      else {
        const fullPath: string = path.resolve(args.options.logoPath);
        if (this.verbose) {
          await logger.logToStderr(`Setting group logo ${fullPath}...`);
        }

        const requestOptionsPhoto: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/photo/$value`,
          headers: {
            'content-type': this.getImageContentType(fullPath)
          },
          data: fs.readFileSync(fullPath)
        };

        await this.setGroupLogo(requestOptionsPhoto, EntraM365GroupAddCommand.numRepeat, logger);
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

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserIds(logger: Logger, users: string | undefined): Promise<string[]> {
    if (!users) {
      if (this.debug) {
        await logger.logToStderr('No users to validate, skipping.');
      }
      return [];
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving user information.');
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
        await setTimeout(this.pollingInterval * (EntraM365GroupAddCommand.numRepeat - retryLeft));
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

export default new EntraM365GroupAddCommand();