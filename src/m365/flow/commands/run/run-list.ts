import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { AzmgmtItemsListCommand } from '../../../base/AzmgmtItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  flowName: string;
}

class FlowRunListCommand extends AzmgmtItemsListCommand<{ name: string, startTime: string, status: string, properties: { startTime: string, status: string } }> {
  public get name(): string {
    return commands.RUN_LIST;
  }

  public get description(): string {
    return 'Lists runs of the specified Microsoft Flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'startTime', 'status'];
  }

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --flowName <flowName>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of runs for Microsoft Flow ${args.options.flowName}...`);
    }

    const url: string = `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environmentName)}/flows/${encodeURIComponent(args.options.flowName)}/runs?api-version=2016-11-01`;

    this
      .getAllItems(url, logger, true)
      .then((): void => {
        if (this.items.length > 0) {
          this.items.forEach(i => {
            i.startTime = i.properties.startTime;
            i.status = i.properties.status;
          });

          logger.log(this.items);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No runs found');
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new FlowRunListCommand();