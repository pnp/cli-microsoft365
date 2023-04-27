import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  displayName?: string;
  tags?: string;
}

class AadSpListCommand extends GraphCommand {
  public get name(): string {
    return commands.SP_LIST;
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'displayName', 'tags'];
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
        tags: typeof args.options.tags !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--displayName [displayName]'
      },
      {
        option: '--tags [tag]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving service principal information...`);
    }

    try {
      let requestUrl: string = `${this.resource}/v1.0/servicePrincipals`;
      const filter: string[] = [];
      if (args.options.tags) {
        const tags: string[] = [];
        args.options.tags.split(',').forEach(t => {
          tags.push(`tags/any(t:t eq '${t}')`);
        });
        filter.push(`(${tags.join(' or ')})`);
      }

      if (args.options.displayName) {
        filter.push(`(displayName eq '${args.options.displayName}')`);
      }

      if (filter.length > 0) {
        requestUrl += `?$filter=${filter.join(' and ')}`;
      }

      logger.log(requestUrl);

      const res = await odata.getAllItems(requestUrl);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadSpListCommand();