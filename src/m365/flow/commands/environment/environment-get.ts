import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';
import * as FlowEnvironmentListCommand from './environment-list';

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
      const options: any = {
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const output = await Cli.executeCommandWithOutput(FlowEnvironmentListCommand as Command, { options: { ...options, _: [] } });
      const flowEnvironmentListOutput = JSON.parse(output.stdout);

      const flowItem: any = flowEnvironmentListOutput.filter(((flow: any) => args.options.name ? flow.name === args.options.name : flow.properties.isDefault === true))[0];

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