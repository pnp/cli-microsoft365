import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
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
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name}...`);
    }

    try {
      const response = await odata.getAllItems<FlowEnvironmentDetails>(`${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`);
      const flowItem: FlowEnvironmentDetails = response.filter(((flow: any) => args.options.name ? flow.name === args.options.name : flow.properties.isDefault === true))[0];

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