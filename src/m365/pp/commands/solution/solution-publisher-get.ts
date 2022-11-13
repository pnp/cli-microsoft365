import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { validation } from '../../../../utils/validation';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import { Publisher } from './Solution';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environment: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpSolutionPublisherGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_GET;
  }

  public get description(): string {
    return 'Get information about the specified publisher in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['publisherid', 'uniquename', 'friendlyname'];
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
        option: '-a, --asAdmin'
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
      logger.logToStderr(`Retrieving a specific publisher '${args.options.id || args.options.name}'...`);
    }

    const res = await this.getSolutionPublisher(args);
    logger.log(res);
  }

  private async getSolutionPublisher(args: CommandArgs): Promise<any> {
    const requestOptions: AxiosRequestConfig = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      if (args.options.id) {
        requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers(${args.options.id})?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;

        const result = await request.get<Publisher>(requestOptions);
        return result;
      }

      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers?$filter=friendlyname eq \'${args.options.name}\'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;
      const result = await request.get<{ value: Publisher[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `The specified publisher '${args.options.name}' does not exist.`;
      }

      return result.value[0];
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpSolutionPublisherGetCommand();