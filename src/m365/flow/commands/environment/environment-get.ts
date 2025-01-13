import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
import { FlowEnvironmentDetails } from './FlowEnvironmentDetails.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
}

class FlowEnvironmentGetCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow environment';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTelemetry();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name ?? ''}...`);
    }

    let requestUrl = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/`;

    if (args.options.name) {
      requestUrl += `${formatting.encodeQueryParameter(args.options.name)}`;
    }
    else {
      requestUrl += `~default`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const flowItem = await request.get<FlowEnvironmentDetails>(requestOptions);

      if (args.options.output !== 'json') {
        flowItem.displayName = flowItem.properties.displayName;
        flowItem.provisioningState = flowItem.properties.provisioningState;
        flowItem.environmentSku = flowItem.properties.environmentSku;
        flowItem.azureRegionHint = flowItem.properties.azureRegionHint;
        flowItem.isDefault = flowItem.properties.isDefault;
      }

      await logger.log(flowItem);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new FlowEnvironmentGetCommand();