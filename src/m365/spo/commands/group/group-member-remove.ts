import { z } from 'zod';
import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand, { Options as EntraUserGetCommandOptions } from '../../../entra/commands/user/user-get.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import SpoGroupMemberListCommand, { Options as SpoGroupMemberListCommandOptions } from './group-member-list.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  groupId: z.number().optional(),
  groupName: z.string().optional(),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `'${e.input}' is not a valid userName.`
    })
    .optional(),
  email: z.string()
    .refine(email => validation.isValidUserPrincipalName(email), {
      error: e => `'${e.input}' is not a valid email.`
    })
    .optional(),
  userId: z.number().optional(),
  entraGroupId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID.`
    })
    .optional(),
  entraGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupMemberRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified member from a SharePoint group';
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
      .refine(o => [o.userName, o.email, o.userId, o.entraGroupId, o.entraGroupName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'userName', 'email', 'userId', 'entraGroupId', or 'entraGroupName', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['userName', 'email', 'userId', 'entraGroupId', 'entraGroupName']
        }
      });
  }

  private async getUserName(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.userName) {
      return args.options.userName;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about the user ${args.options.email}`);
    }

    const options: EntraUserGetCommandOptions = {
      email: args.options.email,
      output: 'json',
      debug: args.options.debug as boolean,
      verbose: args.options.verbose as boolean
    };

    const userGetOutput: CommandOutput = await cli.executeCommandWithOutput(entraUserGetCommand as Command, { options: { ...options, _: [] } });
    const userOutput = JSON.parse(userGetOutput.stdout);
    return userOutput.userPrincipalName;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      if (this.debug) {
        await logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
      }
      await this.removeUserfromSPGroup(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user ${args.options.userName || args.options.userId || args.options.email || args.options.entraGroupId || args.options.entraGroupName} from the SharePoint group?` });

      if (result) {
        await this.removeUserfromSPGroup(logger, args);
      }
    }
  }

  private async removeUserfromSPGroup(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing User ${args.options.userName || args.options.email || args.options.userId || args.options.entraGroupId || args.options.entraGroupName} from Group: ${args.options.groupId || args.options.groupName}`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${args.options.groupId}')`
      : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}`;

    if (args.options.userId) {
      requestUrl += `/users/removeById(${args.options.userId})`;
    }
    else if (args.options.userName || args.options.email) {
      const userName: string = await this.getUserName(logger, args);
      const loginName: string = `i:0#.f|membership|${userName}`;
      requestUrl += `/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`;
    }
    else {
      const entraGroupId = await this.getGroupId(args);
      requestUrl += `/users/RemoveById(${entraGroupId})`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const options: SpoGroupMemberListCommandOptions = {
      webUrl: args.options.webUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    if (args.options.groupId) {
      options.groupId = args.options.groupId;
    }
    else {
      options.groupName = args.options.groupName;
    }

    const output = await cli.executeCommandWithOutput(SpoGroupMemberListCommand as Command, { options: { ...options, _: [] } });
    const getGroupMemberListOutput = JSON.parse(output.stdout);

    let foundGroups: any;

    if (args.options.entraGroupId) {
      foundGroups = getGroupMemberListOutput.filter((x: any) => { return x.LoginName.indexOf(args.options.entraGroupId) > -1 && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }
    else {
      foundGroups = getGroupMemberListOutput.filter((x: any) => { return x.Title === args.options.entraGroupName && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }

    if (foundGroups.length === 0) {
      throw `The Microsoft Entra group ${args.options.entraGroupId || args.options.entraGroupName} is not found in SharePoint group ${args.options.groupId || args.options.groupName}`;
    }

    return foundGroups[0].Id;
  }
}

export default new SpoGroupMemberRemoveCommand();