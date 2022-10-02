import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import { AzmgmtItemsListCommand } from '../../base/AzmgmtItemsListCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin: boolean;
}

class FlowListCommand extends AzmgmtItemsListCommand<{ name: string, displayName: string, properties: { displayName: string } }> {
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flows in the given environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: args.options.asAdmin === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url: string = `${this.resource}providers/Microsoft.ProcessSimple${args.options.asAdmin ? '/scopes/admin' : ''}/environments/${encodeURIComponent(args.options.environment)}/flows?api-version=2016-11-01`;

    try {
      await this.getAllItems(url, logger, true);

      if (this.items.length > 0) {
        this.items.forEach(i => {
          i.displayName = i.properties.displayName;
        });

        logger.log(this.items);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No Flows found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowListCommand();