import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { validation } from '../../../../utils/validation';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import { Publisher, Solution } from './Solution';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
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
    this.#initOptionSets();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin,
        name: typeof args.options.name !== 'undefined',
        id: typeof args.options.id !== 'undefined'
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
      ['id', 'name']
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
      logger.logToStderr(`Retrieving a specific solutions '${args.options.id ? args.options.id : args.options.name}' for which the user is an admin...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);
      const res: Solution = await this.getSolution(dynamicsApiUrl, args.options);
      if (!args.options.output || args.options.output === 'json') {
        logger.log(res);
      }
      else {
        //converted to text friendly output
        logger.log({
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
    const requestOptions: AxiosRequestConfig = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/solutions(${options.id})?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;

      const r: Solution = await request.get<Solution>(requestOptions);
      return r;
    }

    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq \'${options.name}\'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;
    const r = await request.get<{ value: Solution[] }>(requestOptions);
    if (!r.value[0]) {
      throw `The specified solution '${options.name}' does not exist.`;
    }
    return r.value[0];
  }
}

module.exports = new PpSolutionGetCommand();