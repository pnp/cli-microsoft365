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

interface Options extends GlobalOptions {
  environmentName: string;
  newName: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpCardCloneCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CARD_CLONE;
  }

  public get description(): string {
    return 'Clones a specific Microsoft Power Platform card in the specified Power Platform environment.';
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
        option: '--newName <newName>'
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
      await logger.logToStderr(`Cloning a card from '${args.options.id || args.options.name}'...`);
    }

    const res = await this.cloneCard(args, logger);
    await logger.log(res);
  }

  private async getCardId(args: CommandArgs, dynamicsApiUrl: string, logger: Logger): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const card = await powerPlatform.getCardByName(dynamicsApiUrl, args.options.name!, logger, this.verbose);
    return card.cardid;
  }

  private async cloneCard(args: CommandArgs, logger: Logger): Promise<any> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const cardId = await this.getCardId(args, dynamicsApiUrl, logger);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/CardCreateClone`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          CardId: cardId,
          CardName: args.options.newName
        }
      };

      const response = await request.post(requestOptions);
      return response;
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpCardCloneCommand();