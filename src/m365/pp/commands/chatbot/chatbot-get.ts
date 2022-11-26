import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { AxiosRequestConfig } from 'axios';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environment: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpChatbotGetCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.CHATBOT_GET;
  }

  public get description(): string {
    return 'Get specific Microsoft Power Platform chatbot in the specified Power Platform environment.';
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

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving a chatbot '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const res = await this.getChatbot(dynamicsApiUrl, args.options);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getChatbot(dynamicsApiUrl: string, options: Options): Promise<any> {
    const requestOptions: AxiosRequestConfig = {
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

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/bots?$filter=name eq '${options.name}'`;
    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length > 1) {
      throw `Multiple chatbots with name '${options.name}' found: ${result.value.map(x => x.botid).join(',')}`;
    }

    if (result.value.length === 0) {
      throw `The specified chatbot '${options.name}' does not exist.`;
    }

    return result.value[0];
  }
}

module.exports = new PpChatbotGetCommand();