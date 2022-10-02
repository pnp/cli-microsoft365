import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class PaEnvironmentGetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power Apps environment';
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
        option: '-n, --name <name>'
      }
    );
  }
  
  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Power Apps environment ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/providers/Microsoft.PowerApps/environments/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);

      res.displayName = res.properties.displayName;
      res.provisioningState = res.properties.provisioningState;
      res.environmentSku = res.properties.environmentSku;
      res.azureRegionHint = res.properties.azureRegionHint;
      res.isDefault = res.properties.isDefault;

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaEnvironmentGetCommand();