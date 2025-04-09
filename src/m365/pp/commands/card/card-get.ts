import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpCardGetCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CARD_GET;
  }

  public get description(): string {
    return 'Get specific Microsoft Power Platform card in the specified Power Platform environment.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
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
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving a card '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const res = await this.getCard(dynamicsApiUrl, args.options, logger);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCard(dynamicsApiUrl: string, options: Options, logger: Logger): Promise<any> {
    if (options.name) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving the card with name '${options.name}'`);
      }

      return await powerPlatform.getCardByName(dynamicsApiUrl, options.name!);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving the card with id '${options.id}'`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.1/cards(${options.id})`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<any>(requestOptions);
  }
}

export default new PpCardGetCommand();