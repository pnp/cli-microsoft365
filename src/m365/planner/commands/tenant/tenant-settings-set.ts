import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PlannerCommand from '../../../base/PlannerCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  isPlannerAllowed: z.boolean().optional(),
  allowCalendarSharing: z.boolean().optional(),
  allowTenantMoveWithDataLoss: z.boolean().optional(),
  allowTenantMoveWithDataMigration: z.boolean().optional(),
  allowRosterCreation: z.boolean().optional(),
  allowPlannerMobilePushNotifications: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTenantSettingsSetCommand extends PlannerCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets Microsoft Planner configuration of the tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => opts.isPlannerAllowed !== undefined || opts.allowCalendarSharing !== undefined || opts.allowTenantMoveWithDataLoss !== undefined || opts.allowTenantMoveWithDataMigration !== undefined || opts.allowRosterCreation !== undefined || opts.allowPlannerMobilePushNotifications !== undefined, {
        message: 'You must specify at least one option',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        prefer: 'return=representation'
      },
      responseType: 'json',
      data: {
        isPlannerAllowed: args.options.isPlannerAllowed,
        allowCalendarSharing: args.options.allowCalendarSharing,
        allowTenantMoveWithDataLoss: args.options.allowTenantMoveWithDataLoss,
        allowTenantMoveWithDataMigration: args.options.allowTenantMoveWithDataMigration,
        allowRosterCreation: args.options.allowRosterCreation,
        allowPlannerMobilePushNotifications: args.options.allowPlannerMobilePushNotifications
      }
    };

    try {
      const result = await request.patch(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerTenantSettingsSetCommand();
