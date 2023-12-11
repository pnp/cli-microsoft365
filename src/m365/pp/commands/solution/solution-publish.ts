import { AxiosRequestConfig } from 'axios';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppSolutionGetCommand, { Options as PpSolutionGetCommandOptions } from './solution-get.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  name?: string;
  asAdmin?: boolean;
  wait?: boolean;
}

interface SolutionComponent {
  msdyn_componentlogicalname: string;
  msdyn_name: string;
}

class PpSolutionPublishCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISH;
  }

  public get description(): string {
    return 'Publishes the components of a solution in a given Power Platform environment';
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
        option: '--wait'
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
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const solutionId = await this.getSolutionId(args, logger);
      const solutionComponents = await this.getSolutionComponents(dynamicsApiUrl, solutionId, logger);
      const parameterXml = await this.buildXmlRequestObject(solutionComponents, logger);

      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.0/PublishXml`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          ParameterXml: parameterXml
        }
      };

      if (this.verbose) {
        await logger.logToStderr(`Publishing the solution '${args.options.id || args.options.name}'...`);
      }

      if (args.options.wait) {
        await request.post(requestOptions);
      }
      else {
        request.post(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSolutionComponents(dynamicsApiUrl: string, solutionId: string, logger: Logger): Promise<SolutionComponent[]> {
    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${solutionId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (this.verbose) {
      await logger.logToStderr(`Retrieving solution components`);
    }

    const response = await request.get<{ value: SolutionComponent[] }>(requestOptions);

    return response.value;
  }

  private async getSolutionId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving solutionId`);
    }

    const options: PpSolutionGetCommandOptions = {
      environmentName: args.options.environmentName,
      name: args.options.name,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(ppSolutionGetCommand as Command, { options: { ...options, _: [] } });
    const getSolutionOutput = JSON.parse(output.stdout);
    return getSolutionOutput.solutionid;
  }

  private async buildXmlRequestObject(solutionComponents: SolutionComponent[], logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Building the XML request object...`);
    }
    const result = solutionComponents.reduce(function (r, a) {
      const key = a.msdyn_componentlogicalname.slice(-1) === 'y' ?
        a.msdyn_componentlogicalname.substring(0, a.msdyn_componentlogicalname.length - 1) + 'ies' :
        a.msdyn_componentlogicalname + 's';
      r[key] = r[key] || [];

      r[key].push({ [a.msdyn_componentlogicalname]: a.msdyn_name });
      return r;
    }, Object.create(null));

    return `<importexportxml>${formatting.objectToXml(result)}</importexportxml>`;
  }
}

export default new PpSolutionPublishCommand();