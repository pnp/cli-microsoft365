import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
import { FlowEnvironmentDetails } from './FlowEnvironmentDetails.js';

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

class FlowEnvironmentGetCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow environment';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.name !== !options.default, {
        message: `Specify either name or default, but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name ?? 'default'}...`);
    }

    let requestUrl = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/`;
    requestUrl += args.options.default ? '~default' : formatting.encodeQueryParameter(args.options.name!);

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const flowItem = await request.get<FlowEnvironmentDetails>(requestOptions);

      if (args.options.output !== 'json') {
        flowItem.displayName = flowItem.properties.displayName;
        flowItem.provisioningState = flowItem.properties.provisioningState;
        flowItem.environmentSku = flowItem.properties.environmentSku;
        flowItem.azureRegionHint = flowItem.properties.azureRegionHint;
        flowItem.isDefault = flowItem.properties.isDefault;
      }

      await logger.log(flowItem);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowEnvironmentGetCommand();