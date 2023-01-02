import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { Cli } from '../../../../cli/Cli';
import { Options as PpChatbotGetCommandOptions } from './chatbot-get';
import * as PpChatbotGetCommand from './chatbot-get';
import Command from '../../../../Command';

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
        confirm: !!args.options.confirm
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
        option: '--asAdmin'
      },
      {
        option: '--confirm'
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
      logger.logToStderr(`Removing chatbot '${args.options.id || args.options.name}'...`);
    }

    if (args.options.confirm) {
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

  private async getChatbotId(args: CommandArgs): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const options: PpChatbotGetCommandOptions = {
      environment: args.options.environment,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(PpChatbotGetCommand as Command, { options: { ...options, _: [] } });
    const getBotOutput = JSON.parse(output.stdout);
    return getBotOutput.botid;
  }

  private async deleteChatbot(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const botId = await this.getChatbotId(args);
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

module.exports = new PpChatbotRemoveCommand();