import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appName: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID for appName.'
    }),
  asAdmin: z.boolean().optional(),
  environmentName: z.string().optional().alias('e'),
  roleName: z.enum(['Owner', 'CanEdit', 'CanView']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppPermissionListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists all permissions of a Power Apps app';
  }

  public defaultProperties(): string[] | undefined {
    return ['roleName', 'principalId', 'principalType'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.asAdmin || opts.environmentName, {
        message: 'Specifying the environmentName is required when using asAdmin.'
      })
      .refine(opts => !opts.environmentName || opts.asAdmin, {
        message: 'Specifying environmentName is only allowed when using asAdmin.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving permissions for app ${args.options.appName}${args.options.roleName !== undefined ? ` with role name ${args.options.roleName}` : ''}`);
    }

    const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environmentName ? '/environments/' + formatting.encodeQueryParameter(args.options.environmentName) : ''}/apps/${args.options.appName}/permissions?api-version=2022-11-01`;

    try {
      let permissions = await odata.getAllItems<{ principalType: string, principalId: string, roleName: string, properties: { roleName: string, principal: { id: string, type: string } } }>(url);

      if (args.options.roleName) {
        permissions = permissions.filter(permission => permission.properties.roleName === args.options.roleName);
      }

      if (args.options.output !== 'json') {
        permissions.forEach(permission => {
          permission.roleName = permission.properties.roleName;
          permission.principalId = permission.properties.principal.id;
          permission.principalType = permission.properties.principal.type;
        });
      }

      await logger.log(permissions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaAppPermissionListCommand();