import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import { validation } from '../../../../utils/validation';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import { Options as PpSolutionGetCommandOptions } from './solution-get';
import * as PpSolutionGetCommand from './solution-get';
import Command from '../../../../Command';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environment: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
  wait?: boolean;
}

class PpSolutionPublishCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISH;
  }

  public get description(): string {
    return 'Publishes a specific solution in a given environment.';
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
        wait: !!args.options.wait
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
      },
      {
        option: '--wait'
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
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Publishes a specific solution '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);
      this.getSolutionComponents(args, dynamicsApiUrl);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSolutionComponents(args: CommandArgs, dynamicsApiUrl: string): Promise<any> {
    const solutionId = await this.getSolutionId(args);
    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${solutionId})&$orderby=msdyn_displayname asc&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };


    const response = await request.get<any>(requestOptions);
    Cli.log(response);

  }

  private async getSolutionId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const options: PpSolutionGetCommandOptions = {
      environment: args.options.environment,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(PpSolutionGetCommand as Command, { options: { ...options, _: [] } });
    const getSolutionOutput = JSON.parse(output.stdout);
    return getSolutionOutput.solutionid;
  }

}

module.exports = new PpSolutionPublishCommand();