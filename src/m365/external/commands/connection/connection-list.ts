import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import searchCommands from '../../commands.js';
import commands from '../../commands.js';

class ExternalConnectionListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists external connections defined in the Microsoft Search';
  }

  public alias(): string[] | undefined {
    return [searchCommands.EXTERNALCONNECTION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'state'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const connections = await odata.getAllItems(`${this.resource}/v1.0/external/connections`);
      await logger.log(connections);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionListCommand();