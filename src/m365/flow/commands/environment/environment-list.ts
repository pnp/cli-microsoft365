import { CommandArgs } from '../../../../Command';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

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
      logger.logToStderr(`Retrieving list of Microsoft Flow environments...`);
    }

    try {
      const res = await odata.getAllItems<{ name: string, displayName: string; properties: { displayName: string } }>(`${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`);

      if (args.options.output !== 'json' && res.length > 0) {
        res.forEach(e => {
          e.displayName = e.properties.displayName;
        });
      }

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowEnvironmentListCommand();