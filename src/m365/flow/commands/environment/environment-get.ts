import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';
import { FlowEnvironmentDetails } from './FlowEnvironmentDetails';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
}

class FlowEnvironmentGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'id', 'location', 'displayName', 'provisioningState', 'environmentSku', 'azureRegionHint', 'isDefault'];
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
      logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name ?? ''}...`);
    }

    let requestUrl = `${this.resource}providers/Microsoft.ProcessSimple/environments/`;

    if (args.options.name) {
      requestUrl += `${formatting.encodeQueryParameter(args.options.name)}`;
    }
    else {
      requestUrl += `~default`;
    }

    const requestOptions: any = {
      url: `${requestUrl}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const flowItem = await request.get<FlowEnvironmentDetails>(requestOptions);
      flowItem.displayName = flowItem.properties.displayName;
      flowItem.provisioningState = flowItem.properties.provisioningState;
      flowItem.environmentSku = flowItem.properties.environmentSku;
      flowItem.azureRegionHint = flowItem.properties.azureRegionHint;
      flowItem.isDefault = flowItem.properties.isDefault;

      logger.log(flowItem);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowEnvironmentGetCommand();