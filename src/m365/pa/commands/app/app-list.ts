import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { AzmgmtItemsListCommand } from '../../../base/AzmgmtItemsListCommand';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions { }

class PaAppListCommand extends AzmgmtItemsListCommand<{ name: string, properties: { displayName: string } }> {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists all Power Apps apps';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const url: string = `${this.resource}providers/Microsoft.PowerApps/apps?api-version=2017-08-01`;

    this
      .getAllItems(url, cmd, true)
      .then((): void => {
        if (this.items.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(this.items);
          }
          else {
            cmd.log(this.items.map(f => {
              return {
                name: f.name,
                displayName: f.properties.displayName
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No apps found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }
}

module.exports = new PaAppListCommand();