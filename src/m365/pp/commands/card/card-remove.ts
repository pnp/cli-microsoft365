import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
  confirm?: boolean;
}

class PpCardRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CARD_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific Microsoft Power Platform card in the specified Power Platform environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'cardid', 'publishdate', 'createdon', 'modifiedon'];
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
        asAdmin: !!args.options.asAdmin,
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-a, --asAdmin'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['id', 'name']
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

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removes a card '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const cardId = await this.getCardId(dynamicsApiUrl, args.options);

      if (args.options.confirm) {
        await this.deleteCard(dynamicsApiUrl, cardId);
      }
      else {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the card?`
        });

        if (result.continue) {
          await this.deleteCard(dynamicsApiUrl, cardId);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCardId(dynamicsApiUrl: string, options: Options): Promise<any> {
    if (options.id) {
      return options.id;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.1/cards?$filter=name eq '${options.name}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified card '${options.name}' does not exist.`;
    }

    if (result.value.length > 1) {
      throw `Multiple cards with name '${options.name}' found: ${result.value.map(x => x.cardid).join(',')}`;
    }

    return result.value[0].cardid;
  }

  private async deleteCard(dynamicsApiUrl: string, id: string): Promise<void> {
    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.1/cards(${id})`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    await request.delete(requestOptions);
  }
}

module.exports = new PpCardRemoveCommand();