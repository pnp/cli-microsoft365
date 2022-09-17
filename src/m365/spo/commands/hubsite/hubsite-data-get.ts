import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  forceRefresh?: boolean;
}

class SpoHubSiteDataGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_DATA_GET;
  }

  public get description(): string {
    return 'Get hub site data for the specified site';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        forceRefresh: args.options.forceRefresh === true
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --forceRefresh'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Retrieving hub site data...');
    }

    const forceRefresh: boolean = args.options.forceRefresh === true;

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/HubSiteData(${forceRefresh})`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      if (res['odata.null'] !== true) {
        logger.log(JSON.parse(res.value));
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`${args.options.webUrl} is not connected to a hub site and is not a hub site itself`);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoHubSiteDataGetCommand();