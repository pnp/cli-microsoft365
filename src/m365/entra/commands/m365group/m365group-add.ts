import { Group, User } from '@microsoft/microsoft-graph-types';
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

enum Visibility {
  Private = 'Private',
  Public = 'Public',
  HiddenMembership = 'HiddenMembership'
}

const options = globalOptionsZod
  .extend({
    displayName: zod.alias('n', z.string()),
    mailNickname: zod.alias('m', z.string().refine(mailNickname => mailNickname.indexOf(' ') === -1, () => ({
      message: 'The option mailNickname cannot contain spaces.'
    }))),
    description: zod.alias('d', z.string().optional()),
    owners: z.string().refine(owners => validation.isValidUserPrincipalNameArray(owners) === true, owners => ({
      message: `The following owners are invalid: ${validation.isValidUserPrincipalNameArray(owners)}.`
    })).optional(),
    members: z.string().refine(members => validation.isValidUserPrincipalNameArray(members) === true, members => ({
      message: `The following members are invalid: ${validation.isValidUserPrincipalNameArray(members)}.`
    })).optional(),
    visibility: zod.coercedEnum(Visibility).optional(),
    logoPath: zod.alias('l', z.string().optional().refine(logoPath => {
      if (!logoPath) {
        return true; // Skip validation if not provided
      }
      
      const fullPath: string = path.resolve(logoPath);
      
      if (!fs.existsSync(fullPath)) {
        return false;
      }
      
      if (fs.lstatSync(fullPath).isDirectory()) {
        return false;
      }
      
      return true;
    }, logoPath => {
      if (!logoPath) {
        return { message: '' }; // Should never hit this case
      }
      
      const fullPath: string = path.resolve(logoPath);
      
      if (!fs.existsSync(fullPath)) {
        return { message: `File '${fullPath}' not found` };
      }
      
      if (fs.lstatSync(fullPath).isDirectory()) {
        return { message: `Path '${fullPath}' points to a directory` };
      }
      
      return { message: `Invalid file path` };
    })),
    allowMembersToPost: z.boolean().optional(),
    hideGroupInOutlook: z.boolean().optional(),
    subscribeNewGroupMembers: z.boolean().optional(),
    welcomeEmailDisabled: z.boolean().optional()
  })
  .strict();

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

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  constructor() {
    super();

    this.#initTelemetry();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        description: typeof args.options.description !== 'undefined',
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: Group;
    let ownerIds: string[] = [];
    let memberIds: string[] = [];
    const resourceBehaviorOptionsCollection: string[] = [];
    const resolvedVisibility = args.options.visibility || Visibility.Public;

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

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
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