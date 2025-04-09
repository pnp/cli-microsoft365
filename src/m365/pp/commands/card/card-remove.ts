import { cli } from '../../../../cli/cli.js';
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
  id?: string;
  name?: string;
  asAdmin?: boolean;
  force?: boolean;
}

class PpCardRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CARD_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific Microsoft Power Platform card in the specified Power Platform environment.';
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
        force: !!args.options.force
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
      },
      {
        option: '-f, --force'
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
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing card '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteCard(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove card '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.deleteCard(args, logger);
      }
    }
  }

  private async getCardId(args: CommandArgs, dynamicsApiUrl: string, logger: Logger): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving the card with name '${args.options.name}'`);
    }

    const card = await powerPlatform.getCardByName(dynamicsApiUrl, args.options.name!);

    return card.cardid;
  }

  private async deleteCard(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const cardId = await this.getCardId(args, dynamicsApiUrl, logger);

      if (this.verbose) {
        await logger.logToStderr(`Deleting card with Id '${cardId}'`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/cards(${cardId})`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpCardRemoveCommand();