import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().optional().alias('e'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists all Power Apps apps';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.asAdmin || opts.environmentName, {
        message: 'When specifying the asAdmin option the environment option is required as well.'
      })
      .refine(opts => !opts.environmentName || opts.asAdmin, {
        message: 'When specifying the environment option the asAdmin option is required as well.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environmentName ? '/environments/' + formatting.encodeQueryParameter(args.options.environmentName) : ''}/apps?api-version=2017-08-01`;

    try {
      const apps = await odata.getAllItems<{ name: string; displayName: string; properties: { displayName: string } }>(url);

      if (apps.length > 0) {
        apps.forEach(a => {
          a.displayName = a.properties.displayName;
        });
      }
      await logger.log(apps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaAppListCommand();