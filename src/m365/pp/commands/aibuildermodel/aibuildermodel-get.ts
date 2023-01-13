import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { validation } from '../../../../utils/validation';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpAiBuilderModelGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.AIBUILDERMODEL_GET;
  }

  public get description(): string {
    return 'Get an AI builder model in the specified Power Platform environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['msdyn_name', 'msdyn_aimodelid', 'createdon', 'modifiedon'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
    this.#initValidators();
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving an AI builder model '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const res = await this.getAiBuilderModel(dynamicsApiUrl, args.options);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAiBuilderModel(dynamicsApiUrl: string, options: Options): Promise<any> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${options.id})?$filter=iscustomizable/Value eq true`;
      const result = await request.get<any>(requestOptions);
      return result;
    }

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${options.name}' and iscustomizable/Value eq true`;
    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified AI builder model '${options.name}' does not exist.`;
    }

    if (result.value.length > 1) {
      throw `Multiple AI builder models with name '${options.name}' found: ${result.value.map(x => x.msdyn_aimodelid).join(',')}`;
    }

    return result.value[0];
  }
}

module.exports = new PpAiBuilderModelGetCommand();