import { Group, User } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const VisibilityEnum = {
  Private: 'Private',
  Public: 'Public',
  HiddenMembership: 'HiddenMembership'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  displayName: z.string().alias('n'),
  mailNickname: z.string().alias('m'),
  description: z.string().optional().alias('d'),
  owners: z.string()
    .refine(owners => owners.split(',').map(o => o.trim()).every(o => o.indexOf('@') > -1), {
      error: e => `${(e.input as string).split(',').map(o => o.trim()).find(o => o.indexOf('@') < 0)} is not a valid userPrincipalName`
    }).optional(),
  members: z.string()
    .refine(members => members.split(',').map(m => m.trim()).every(m => m.indexOf('@') > -1), {
      error: e => `${(e.input as string).split(',').map(m => m.trim()).find(m => m.indexOf('@') < 0)} is not a valid userPrincipalName`
    }).optional(),
  visibility: zod.coercedEnum(VisibilityEnum).optional(),
  logoPath: z.string()
    .refine(logoPath => {
      const fullPath = path.resolve(logoPath);
      return fs.existsSync(fullPath) && !fs.lstatSync(fullPath).isDirectory();
    }, {
      error: e => {
        const fullPath = path.resolve(e.input as string);
        return fs.existsSync(fullPath) ? `Path '${fullPath}' points to a directory` : `File '${fullPath}' not found`;
      }
    })
    .optional()
    .alias('l'),
  allowMembersToPost: z.boolean().optional(),
  hideGroupInOutlook: z.boolean().optional(),
  subscribeNewGroupMembers: z.boolean().optional(),
  welcomeEmailDisabled: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupAddCommand extends GraphCommand {
  private static numRepeat: number = 15;
  private pollingInterval: number = 500;

  public get name(): string {
    return commands.M365GROUP_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft 365 Group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.mailNickname.indexOf(' ') === -1, {
        error: 'The option mailNickname cannot contain spaces.',
        path: ['mailNickname']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: Group;
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
      const ownerIds = await this.getUserIds(logger, args.options.owners);
      const memberIds = await this.getUserIds(logger, args.options.members);
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

    const promises: Promise<{ value: User[] }>[] = userArr.map(user => {
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
    const userIds = usersRes.map(res => res.value[0]?.id as string);

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
    const extension: string = imagePath.substring(imagePath.lastIndexOf('.')).toLowerCase();

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