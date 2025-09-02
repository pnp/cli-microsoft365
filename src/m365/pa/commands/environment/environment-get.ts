import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string().optional()),
    default: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaEnvironmentGetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power Apps environment';
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
      await logger.logToStderr(`Retrieving information about Microsoft Power Apps environment ${args.options.name || 'default'}...`);
    }

    const environmentName = args.options.default ? '~default' : formatting.encodeQueryParameter(args.options.name!);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/environments/${environmentName}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);

      res.displayName = res.properties.displayName;
      res.provisioningState = res.properties.provisioningState;
      res.environmentSku = res.properties.environmentSku;
      res.azureRegionHint = res.properties.azureRegionHint;
      res.isDefault = res.properties.isDefault;

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaEnvironmentGetCommand();