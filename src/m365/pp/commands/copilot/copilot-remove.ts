import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppCopilotGetCommand, { Options as PpCopilotGetCommandOptions } from './copilot-get.js';

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

class PpCopilotRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.COPILOT_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified copilot';
  }

  public alias(): string[] | undefined {
    return [commands.CHATBOT_REMOVE];
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
    await this.showDeprecationWarning(logger, commands.CHATBOT_REMOVE, commands.COPILOT_REMOVE);
    if (this.verbose) {
      await logger.logToStderr(`Removing copilot '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteCopilot(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove copilot '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.deleteCopilot(args);
      }
    }
  }

  private async getCopilotId(args: CommandArgs): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const options: PpCopilotGetCommandOptions = {
      environmentName: args.options.environmentName,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(ppCopilotGetCommand as Command, { options: { ...options, _: [] } });
    const getBotOutput = JSON.parse(output.stdout);
    return getBotOutput.botid;
  }

  private async deleteCopilot(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const botId = await this.getCopilotId(args);
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

export default new PpCopilotRemoveCommand();