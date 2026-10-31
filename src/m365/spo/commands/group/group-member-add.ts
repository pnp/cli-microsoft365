import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  groupId: z.number().int().positive().optional(),
  groupName: z.string().optional(),
  userNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `Option 'userNames' contains one or more invalid UPNs: ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional(),
  emails: z.string()
    .refine(emails => validation.isValidUserPrincipalNameArray(emails) === true, {
      error: e => `Option 'emails' contains one or more invalid UPNs: ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    })
    .optional(),
  userIds: z.string()
    .refine(ids => validation.isValidPositiveIntegerArray(ids) === true, {
      error: e => `Option 'userIds' contains one or more invalid numbers: ${validation.isValidPositiveIntegerArray(e.input as string)}.`
    })
    .optional(),
  entraGroupIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `Option 'entraGroupIds' contains one or more invalid GUIDs: ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional(),
  entraGroupNames: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupMemberAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds members to a SharePoint Group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.groupId, o.groupName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'groupId' or 'groupName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName']
        }
      })
      .refine(o => [o.userNames, o.emails, o.userIds, o.entraGroupIds, o.entraGroupNames].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'userNames', 'emails', 'userIds', 'entraGroupIds', or 'entraGroupNames', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['userNames', 'emails', 'userIds', 'entraGroupIds', 'entraGroupNames']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const loginNames = await this.getLoginNames(logger, args.options);

      let apiUrl = `${args.options.webUrl}/_api/web/SiteGroups`;
      if (args.options.groupId) {
        apiUrl += `/GetById(${args.options.groupId})`;
      }
      else {
        apiUrl += `/GetByName('${formatting.encodeQueryParameter(args.options.groupName!)}')`;
      }
      apiUrl += '/users';

      if (this.verbose) {
        await logger.logToStderr('Adding members to group...');
      }

      const result = [];
      for (const loginName of loginNames) {
        const requestOptions: CliRequestOptions = {
          url: apiUrl,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            LoginName: loginName
          }
        };
        const response = await request.post(requestOptions);
        result.push(response);
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getLoginNames(logger: Logger, options: Options): Promise<string[]> {
    const loginNames: string[] = [];

    if (options.userNames || options.emails) {
      loginNames.push(...formatting.splitAndTrim(options.userNames || options.emails!).map(u => `i:0#.f|membership|${u}`));
    }
    else if (options.entraGroupIds || options.entraGroupNames) {
      if (this.verbose) {
        await logger.logToStderr(`Resolving ${(options.entraGroupIds || options.entraGroupNames)!.length} group(s)...`);
      }

      const groups: Group[] = [];
      if (options.entraGroupIds) {
        const groupIds = formatting.splitAndTrim(options.entraGroupIds);

        for (const groupId of groupIds) {
          const group = await entraGroup.getGroupById(groupId);
          groups.push(group);
        }
      }
      else {
        const groupNames = formatting.splitAndTrim(options.entraGroupNames!);

        for (const groupName of groupNames) {
          const group = await entraGroup.getGroupByDisplayName(groupName);
          groups.push(group);
        }
      }

      // Check if group is M365 group or security group
      loginNames.push(...groups.map(g => g.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${g.id}` : `c:0t.c|tenant|${g.id}`));
    }
    else if (options.userIds) {
      const userIds = formatting.splitAndTrim(options.userIds);

      if (this.verbose) {
        await logger.logToStderr(`Resolving ${userIds.length} user(s)...`);
      }

      for (const userId of userIds) {
        const loginName = await this.getUserLoginNameById(options.webUrl, parseInt(userId));
        loginNames.push(loginName);
      }
    }

    return loginNames;
  }

  private async getUserLoginNameById(webUrl: string, userId: number): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/SiteUsers/GetById(${userId})?$select=LoginName`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const user = await request.get<{ LoginName: string }>(requestOptions);
    return user.LoginName;
  }
}

export default new SpoGroupMemberAddCommand();