import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { AzmgmtItemsListCommand } from '../../../base/AzmgmtItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions { }

class PaAppListCommand extends AzmgmtItemsListCommand<{ name: string; displayName: string; properties: { displayName: string } }> {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists all Power Apps apps';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const url: string = `${this.resource}providers/Microsoft.PowerApps/apps?api-version=2017-08-01`;

    this
      .getAllItems(url, logger, true)
      .then((): void => {
        if (this.items.length > 0) {
          this.items.forEach(a => {
            a.displayName = a.properties.displayName;
          });

          logger.log(this.items);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No apps found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new PaAppListCommand();