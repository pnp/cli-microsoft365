import { z } from 'zod';
import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand, { Options as EntraUserGetCommandOptions } from '../../../entra/commands/user/user-get.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.number().optional().alias('i'),
  name: z.string().optional().alias('n'),
  newName: z.string().optional(),
  description: z.string().optional(),
  allowMembersEditMembership: z.boolean().optional(),
  onlyAllowMembersViewMembership: z.boolean().optional(),
  allowRequestToJoinLeave: z.boolean().optional(),
  autoAcceptRequestToJoinLeave: z.boolean().optional(),
  requestToJoinLeaveEmailSetting: z.string().optional(),
  ownerEmail: z.string().optional(),
  ownerUserName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupSetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a group in the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.id, o.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(o => [o.ownerEmail, o.ownerUserName].filter(x => x !== undefined).length <= 1, {
        error: `Specify either 'ownerEmail' or 'ownerUserName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['ownerEmail', 'ownerUserName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Setting properties for group ${args.options.id || args.options.name}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/sitegroups/${args.options.id ? `GetById(${args.options.id})` : `GetByName('${args.options.name}')`}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        Title: args.options.newName,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    try {
      await request.patch(requestOptions);
      if (args.options.ownerEmail || args.options.ownerUserName) {
        await this.setGroupOwner(args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async setGroupOwner(options: Options): Promise<void> {
    const ownerId = await this.getOwnerId(options);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/sitegroups/${options.id ? `GetById(${options.id})` : `GetByName('${options.name}')`}/SetUserAsOwner(${ownerId})`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getOwnerId(options: Options): Promise<number> {
    const cmdOptions: EntraUserGetCommandOptions = {
      userName: options.ownerUserName,
      email: options.ownerEmail,
      output: 'json',
      debug: options.debug as boolean,
      verbose: options.verbose as boolean
    };

    const output: CommandOutput = await cli.executeCommandWithOutput(entraUserGetCommand as Command, { options: { ...cmdOptions, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/ensureUser('${getUserOutput.userPrincipalName}')?$select=Id`,
      headers: {
        accept: 'application/json',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.post<{ Id: number }>(requestOptions);
    return response.Id;
  }
}

export default new SpoGroupSetCommand();