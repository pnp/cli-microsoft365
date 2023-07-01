import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
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
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of cards`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const items = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.1/cards?$expand=owninguser($select=azureactivedirectoryobjectid,fullname)`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpCardListCommand();