import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppAiBuilderModelGetCommand, { Options as PpAiBuilderModelGetCommandOptions } from './aibuildermodel-get.js';

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

class PpAiBuilderModelRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.AIBUILDERMODEL_REMOVE;
  }

  public get description(): string {
    return 'Removes an AI builder model in the specified Power Platform environment.';
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
      await logger.logToStderr(`Removing AI builder model '${args.options.id || args.options.name}'...`);
    }

    if (args.options.force) {
      await this.deleteAiBuilderModel(args);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove AI builder model '${args.options.id || args.options.name}'?`);

      if (result) {
        await this.deleteAiBuilderModel(args);
      }
    }
  }

  private async getAiBuilderModelId(args: CommandArgs): Promise<any> {
    if (args.options.id) {
      return args.options.id;
    }

    const options: PpAiBuilderModelGetCommandOptions = {
      environmentName: args.options.environmentName,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(ppAiBuilderModelGetCommand as Command, { options: { ...options, _: [] } });
    const getAiBuilderModelOutput = JSON.parse(output.stdout);
    return getAiBuilderModelOutput.msdyn_aimodelid;
  }

  private async deleteAiBuilderModel(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const aiBuilderModelId = await this.getAiBuilderModelId(args);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${aiBuilderModelId})`,
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

export default new PpAiBuilderModelRemoveCommand();