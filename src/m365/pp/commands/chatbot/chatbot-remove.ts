import { Cli } from '../../../../cli/Cli.js';
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

class PpChatbotRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CHATBOT_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified chatbot';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing chatbot '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteChatbot(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove chatbot '${args.options.id || args.options.name}'?`
      });

      if (result.continue) {
        await this.deleteChatbot(args);
      }
    }
  }

  private async getChatbotId(args: CommandArgs, dynamicsApiUrl: string): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const chatbot = await powerPlatform.getChatbotByName(dynamicsApiUrl, args.options.name!);
    return chatbot.botid;
  }

  private async deleteChatbot(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const botId = await this.getChatbotId(args, dynamicsApiUrl);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/bots(${botId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`,
        headers: {
          accept: 'application/json',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpChatbotRemoveCommand();