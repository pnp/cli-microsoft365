import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  appName: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID for appName.'
    }),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID for userId.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value is not a valid UPN for userName.'
    })
    .optional(),
  roleForOldAppOwner: z.enum(['CanView', 'CanEdit']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppOwnerSetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_OWNER_SET;
  }

  public get description(): string {
    return 'Sets a new owner for a Power Apps app';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.userId, opts.userName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'userId' or 'userName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Setting new owner ${args.options.userId || args.options.userName} for Power Apps app ${args.options.appName}...`);
    }
    try {
      const userId = await this.getUserId(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.appName}/modifyAppOwner?api-version=2022-11-01`,
        headers: {
          accept: 'application/json',
          'Content-Type': 'application/json'
        },
        responseType: 'json',
        data: {
          roleForOldAppOwner: args.options.roleForOldAppOwner,
          newAppOwner: userId
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(options: Options): Promise<string> {
    if (options.userId) {
      return options.userId;
    }

    return entraUser.getUserIdByUpn(options.userName!);
  }
}

export default new PaAppOwnerSetCommand();