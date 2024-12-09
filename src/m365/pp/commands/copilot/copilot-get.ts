import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpCopilotGetCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.COPILOT_GET;
  }

  public get description(): string {
    return 'Get information about the specified copilot';
  }

  public alias(): string[] | undefined {
    return [commands.CHATBOT_GET];
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'botid', 'publishedon', 'createdon', 'modifiedon'];
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
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, commands.CHATBOT_GET, commands.COPILOT_GET);
    if (this.verbose) {
      await logger.logToStderr(`Retrieving copilot '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const res = await this.getCopilot(dynamicsApiUrl, args.options);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCopilot(dynamicsApiUrl: string, options: Options): Promise<any> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/bots(${options.id})`;
      const result = await request.get<any>(requestOptions);
      return result;
    }

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(options.name!)}'`;
    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('botid', result.value);
      return await cli.handleMultipleResultsFound(`Multiple copilots with name '${options.name}' found.`, resultAsKeyValuePair);
    }

    if (result.value.length === 0) {
      throw `The specified copilot '${options.name}' does not exist.`;
    }

    return result.value[0];
  }
}

export default new PpCopilotGetCommand();