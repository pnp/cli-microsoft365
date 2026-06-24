import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import Auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
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
      message: 'The value is not a valid user principal name (UPN) for userName.'
    })
    .optional(),
  groupId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID for groupId.'
    })
    .optional(),
  groupName: z.string().optional(),
  tenant: z.boolean().optional(),
  asAdmin: z.boolean().optional(),
  environmentName: z.string().optional().alias('e'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppPermissionRemoveCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_PERMISSION_REMOVE;
  }

  public get description(): string {
    return 'Removes permissions to a Power Apps app';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.userId, opts.userName, opts.groupId, opts.groupName, opts.tenant].filter(x => x !== undefined).length === 1, {
        error: `Specify exactly one of the following options: 'userId', 'userName', 'groupId', 'groupName' or 'tenant'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName', 'tenant']
        }
      })
      .refine(opts => !opts.environmentName || opts.asAdmin, {
        message: 'Specifying environmentName is only allowed when using asAdmin.'
      })
      .refine(opts => !opts.asAdmin || opts.environmentName, {
        message: 'Specifying asAdmin is only allowed when using environmentName.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        await this.removeAppPermission(logger, args.options);
      }
      else {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the permissions of '${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || (args.options.tenant && 'everyone')}' from the Power App '${args.options.appName}'?` });

        if (result) {
          await this.removeAppPermission(logger, args.options);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeAppPermission(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing permissions for '${options.userId || options.userName || options.groupId || options.groupName || (options.tenant && 'everyone')}' for the Power Apps app ${options.appName}...`);
    }

    const principalId: string = await this.getPrincipalId(options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/${options.asAdmin ? `scopes/admin/environments/${options.environmentName}/` : ''}apps/${options.appName}/modifyPermissions?api-version=2022-11-01`,
      headers: {
        accept: 'application/json'
      },
      data: {
        delete: [
          {
            id: principalId
          }
        ]
      },
      responseType: 'json'
    };

    await request.post<any>(requestOptions);
  }

  private async getPrincipalId(options: Options): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }
    if (options.userId) {
      return options.userId;
    }
    if (options.groupName) {
      const group = await entraGroup.getGroupByDisplayName(options.groupName);
      return group.id!;
    }
    if (options.userName) {
      const userId = await entraUser.getUserIdByUpn(options.userName);
      return userId;
    }

    return `tenant-${accessToken.getTenantIdFromAccessToken(Auth.connection.accessTokens[Auth.defaultResource].accessToken)}`;
  }
}

export default new PaAppPermissionRemoveCommand();