import { CommandArgs } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';

const options = globalOptionsZod.strict();

class FlowEnvironmentListCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of Microsoft Flow environments...`);
    }

    try {
      const res = await odata.getAllItems<{ name: string, displayName: string; properties: { displayName: string } }>(`${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`);

      if (res.length > 0) {
        if (args.options.output !== 'json') {
          res.forEach(e => {
            e.displayName = e.properties.displayName;
          });
        }
      }
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowEnvironmentListCommand();