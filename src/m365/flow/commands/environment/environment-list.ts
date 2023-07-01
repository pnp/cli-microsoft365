import { CommandArgs } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import AzmgmtCommand from '../../../base/AzmgmtCommand.js';
import commands from '../../commands.js';

class FlowEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of Microsoft Flow environments...`);
    }

    try {
      const res = await odata.getAllItems<{ name: string, displayName: string; properties: { displayName: string } }>(`${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`);

      if (args.options.output !== 'json' && res.length > 0) {
        res.forEach(e => {
          e.displayName = e.properties.displayName;
        });

        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowEnvironmentListCommand();