import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  flowName: z.uuid(),
  environmentName: z.string().alias('e'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowRunCancelCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.RUN_CANCEL;
  }

  public get description(): string {
    return 'Cancels a specific run of the specified Microsoft Flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.log(`Cancelling run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const cancelFlow = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}/cancel?api-version=2016-11-01`,
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
    };

    if (args.options.force) {
      await cancelFlow();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to cancel the flow run ${args.options.name}?` });

      if (result) {
        await cancelFlow();
      }
    }
  }
}

export default new FlowRunCancelCommand();