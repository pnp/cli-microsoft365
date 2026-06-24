import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import Auth from '../../../../Auth.js';
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
  roleName: z.enum(['CanEdit', 'CanView']),
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
  environmentName: z.string().optional().alias('e'),
  sendInvitationMail: z.boolean().optional(),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppPermissionEnsureCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_PERMISSION_ENSURE;
  }

  public get description(): string {
    return 'Assigns/updates permissions to a Power Apps app';
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
      })
      .refine(opts => !opts.tenant || opts.roleName === 'CanView', {
        message: 'Sharing with the entire tenant is only supported with CanView role.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Assigning/updating permissions for '${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || (args.options.tenant && 'everyone')}' to the Power Apps app '${args.options.appName}'...`);
    }

    try {
      const principalId = await this.getPrincipalId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.PowerApps/${args.options.asAdmin ? `scopes/admin/environments/${args.options.environmentName}/` : ''}apps/${args.options.appName}/modifyPermissions?api-version=2022-11-01`,
        headers: {
          accept: 'application/json'
        },
        data: {
          put: [
            {
              properties: {
                principal: {
                  id: principalId,
                  type: this.getPrincipalType(args.options)
                },
                NotifyShareTargetOption: args.options.sendInvitationMail ? 'Notify' : 'DoNotNotify',
                roleName: args.options.roleName
              }
            }
          ]
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getPrincipalType(options: Options): string {
    if (options.userId || options.userName) {
      return 'User';
    }
    if (options.groupId || options.groupName) {
      return 'Group';
    }

    return 'Tenant';
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

    return accessToken.getTenantIdFromAccessToken(Auth.connection.accessTokens[Auth.defaultResource].accessToken);
  }
}

export default new PaAppPermissionEnsureCommand();