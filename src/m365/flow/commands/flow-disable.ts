import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  environmentName: z.string().alias('e'),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowDisableCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.DISABLE;
  }

  public get description(): string {
    return 'Disables specified Microsoft Flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Disables Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/stop?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowDisableCommand();