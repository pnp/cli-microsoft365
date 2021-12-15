import { Logger } from '../../../../cli';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class SearchExternalConnectionListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.EXTERNALCONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists all external connections';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/external/connections`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SearchExternalConnectionListCommand();
