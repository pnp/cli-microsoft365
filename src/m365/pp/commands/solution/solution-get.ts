import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { Publisher, Solution } from './Solution.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
}

class PpSolutionGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_GET;
  }

  public get description(): string {
    return 'Gets a specific solution in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['uniquename', 'version', 'publisher'];
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
    if (this.verbose) {
      await logger.logToStderr(`Retrieving a specific solution '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const res = await this.getSolution(dynamicsApiUrl, args.options);

      if (!args.options.output || !Cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res);
      }
      else {
        // Converted to text friendly output
        await logger.log({
          uniquename: res.uniquename,
          version: res.version,
          publisher: (res.publisherid as Publisher).friendlyname
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSolution(dynamicsApiUrl: string, options: Options): Promise<Solution> {
    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/solutions(${options.id})?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;

      const result = await request.get<Solution>(requestOptions);
      return result;
    }

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq \'${options.name}\'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;
    const result = await request.get<{ value: Solution[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified solution '${options.name}' does not exist.`;
    }

    return result.value[0];
  }
}

export default new PpSolutionGetCommand();