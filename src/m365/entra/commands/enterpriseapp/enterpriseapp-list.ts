import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName?: string;
  tag?: string;
}

class EntraEnterpriseAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_LIST;
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'displayName', 'tag'];
  }

  public get description(): string {
    return 'Lists the enterprise applications (or service principals) in Entra ID';
  }

  public alias(): string[] | undefined {
    return [aadCommands.SP_LIST, commands.SP_LIST];
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
    this.showDeprecationWarning(logger, aadCommands.SP_LIST, commands.SP_LIST);

    if (this.verbose) {
      await logger.logToStderr(`Retrieving enterprise application information...`);
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

export default new EntraEnterpriseAppListCommand();