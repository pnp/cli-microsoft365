import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin?: boolean;
}

class PpCardListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.CARD_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform cards in the specified Power Platform environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'cardid', 'publishdate', 'createdon', 'modifiedon'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '-a, --asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of cards`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const items = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.1/cards?$expand=owninguser($select=azureactivedirectoryobjectid,fullname)`);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpCardListCommand();