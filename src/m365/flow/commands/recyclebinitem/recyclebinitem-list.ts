import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    environmentName: zod.alias('e', z.string())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowRecycleBinItemListCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists all soft-deleted Power Automate flows within an environment';
  }

  public defaultProperties(): string[] {
    return ['name', 'displayName'];
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Getting list of soft-deleted flows in environment ${args.options.environmentName}...`);
      }

      const flows = await odata.getAllItems<any>(`${this.resource}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`);
      const deletedFlows = flows.filter(flow => flow.properties.state === 'Deleted');

      if (cli.shouldTrimOutput(args.options.output)) {
        deletedFlows.forEach(flow => {
          flow.displayName = flow.properties.displayName;
        });
      }

      await logger.log(deletedFlows);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowRecycleBinItemListCommand();