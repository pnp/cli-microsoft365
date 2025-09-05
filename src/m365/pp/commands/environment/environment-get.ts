import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string().optional()),
    default: z.boolean().optional(),
    asAdmin: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpEnvironmentGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Power Platform environment';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !!options.name !== !!options.default, {
        message: `Specify either name or default, but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving environment: ${args.options.name || 'default'}`);
    }

    let url: string = `${this.resource}/providers/Microsoft.BusinessAppPlatform`;
    if (args.options.asAdmin) {
      url += '/scopes/admin';
    }

    const envName = args.options.default ? '~Default' : formatting.encodeQueryParameter(args.options.name!);
    url += `/environments/${envName}?api-version=2020-10-01`;

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    await logger.log(response);
  }
}

export default new PpEnvironmentGetCommand();