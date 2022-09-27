import { Logger } from '../../../../cli';
import GraphCommand from '../../../base/GraphCommand';
import { odata } from '../../../../utils';
import commands from '../../commands';

class SearchExternalConnectionListCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists external connections defined in the Microsoft Search';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'state'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const connections = await odata.getAllItems(`${this.resource}/v1.0/external/connections`);
      logger.log(connections);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SearchExternalConnectionListCommand();