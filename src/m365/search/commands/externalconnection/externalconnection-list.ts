import { Logger } from '../../../../cli';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class SearchExternalConnectionListCommand extends GraphCommand {
  private items: ExternalConnectors.ExternalConnection[] = [];

  public get name(): string {
    return commands.EXTERNALCONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists all external connections';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/external/connections`;

    odata
      .getAllItems<ExternalConnectors.ExternalConnection>(endpoint, logger)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SearchExternalConnectionListCommand();