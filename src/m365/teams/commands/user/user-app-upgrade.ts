import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  name: z.string().optional(),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value must be a valid user principal name.'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsUserAppUpgradeCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades an app in the personal scope of the specified user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(o => o !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => [opts.userId, opts.userName].filter(o => o !== undefined).length === 1, {
        message: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Upgrading app ${args.options.id || args.options.name} for user ${args.options.userId || args.options.userName}`);
      }

      const installedAppId: string = await this.getInstalledAppId(args.options, logger);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName!)}/teamwork/installedApps/${installedAppId}/upgrade`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getInstalledAppId(options: Options, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving app ID`);
    }

    if (options.id) {
      return options.id;
    }

    const installedApps = await odata.getAllItems<{ id: string }>(`${this.resource}/v1.0/users/${formatting.encodeQueryParameter(options.userId || options.userName!)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(options.name!)}'&$select=id`);

    if (installedApps.length === 1) {
      return installedApps[0].id;
    }

    if (installedApps.length === 0) {
      throw `The specified Teams app ${options.name!} does not exist or is not installed for the user`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', installedApps);
    const result: { id: string } = (await cli.handleMultipleResultsFound(`Multiple installed Teams apps with name '${options.name}' found.`, resultAsKeyValuePair)) as { id: string };
    return result.id;
  }
}

export default new TeamsUserAppUpgradeCommand();