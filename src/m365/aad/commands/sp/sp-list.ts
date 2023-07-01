import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  displayName?: string;
  tag?: string;
}

class AadSpListCommand extends GraphCommand {
  public get name(): string {
    return commands.SP_LIST;
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'displayName', 'tag'];
  }

  public get description(): string {
    return 'Lists the service principals in the directory';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        tag: typeof args.options.tag !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--displayName [displayName]'
      },
      {
        option: '--tag [tag]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving service principal information...`);
    }

    try {
      let requestUrl: string = `${this.resource}/v1.0/servicePrincipals`;
      const filter: string[] = [];

      if (args.options.tag) {
        filter.push(`(tags/any(t:t eq '${args.options.tag}'))`);
      }

      if (args.options.displayName) {
        filter.push(`(displayName eq '${args.options.displayName}')`);
      }

      if (filter.length > 0) {
        requestUrl += `?$filter=${filter.join(' and ')}`;
      }

      const res = await odata.getAllItems(requestUrl);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadSpListCommand();