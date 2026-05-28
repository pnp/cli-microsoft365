import { z } from 'zod';
import chalk from 'chalk';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';

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

class FlowRunResubmitCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.RUN_RESUBMIT;
  }

  public get description(): string {
    return 'Resubmits a specific flow run for the specified Microsoft Flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const resubmitFlow = async (): Promise<void> => {
      try {
        const triggerName = await this.getTriggerName(args.options.environmentName, args.options.flowName);

        if (this.debug) {
          await logger.logToStderr(chalk.yellow(`Retrieved trigger: ${triggerName}`));
        }

        const requestOptions: CliRequestOptions = {
          url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/triggers/${formatting.encodeQueryParameter(triggerName)}/histories/${formatting.encodeQueryParameter(args.options.name)}/resubmit?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await resubmitFlow();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to resubmit the flow with run ${args.options.name}?` });

      if (result) {
        await resubmitFlow();
      }
    }
  }

  private async getTriggerName(environment: string, flow: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environment)}/flows/${formatting.encodeQueryParameter(flow)}/triggers?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { name: string; }[]; }>(requestOptions);
    return res.value[0].name;
  }
}

export default new FlowRunResubmitCommand();