import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        res.displayName = res.properties.displayName;
        res.provisioningState = res.properties.provisioningState;
        res.environmentSku = res.properties.environmentSku;
        res.azureRegionHint = res.properties.azureRegionHint;
        res.isDefault = res.properties.isDefault;

        logger.log(res);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowEnvironmentGetCommand();