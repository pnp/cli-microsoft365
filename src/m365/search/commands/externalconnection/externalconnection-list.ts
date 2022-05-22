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

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    odata
      .getAllItems(`${this.resource}/v1.0/external/connections`)
      .then((connections): void => {
        logger.log(connections);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SearchExternalConnectionListCommand();