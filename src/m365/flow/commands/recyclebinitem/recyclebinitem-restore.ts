import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  flowName: z.uuid().alias('n')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowRecycleBinItemRestoreCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores a soft-deleted Power Automate flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Restoring soft-deleted flow ${args.options.flowName} from environment ${args.options.environmentName}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${args.options.flowName}/restore?api-version=2016-11-01`,
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
  }
}

export default new FlowRecycleBinItemRestoreCommand();