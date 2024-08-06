import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    environmentName: zod.alias('e', z.string()),
    flowName: zod.alias('n', z.string()
      .refine(name => validation.isValidGuid(name), name => ({
        message: `'${name}' is not a valid GUID.`
      }))
    )
  })
  .strict();
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

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Restoring soft-deleted flow ${args.options.flowName} from environment ${args.options.environmentName}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${args.options.flowName}/restore?api-version=2016-11-01`,
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