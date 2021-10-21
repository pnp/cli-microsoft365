import { Logger } from '../../../../cli';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
//import ExternalConnection from '../../ExternalConnection';
import GlobalOptions from '../../../../GlobalOptions';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filter?: string;
}

class SearchExternalConnectionListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return commands.EXTERNALCONNECTION_LIST;
  }

  public get description(): string {
    return 'Adds a new External Connection for Microsoft Search';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name'];
  }


  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/external/connections`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => {
        logger.log(err.message);
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }
}

module.exports = new SearchExternalConnectionListCommand();