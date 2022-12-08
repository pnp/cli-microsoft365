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

export interface SolutionComponent {
  msdyn_componentlogicalname: string;
  msdyn_name: string;
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
    if (this.verbose) {
      logger.logToStderr(`Publishes a specific solution '${args.options.id || args.options.name}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const parameterXml = await this.getSolutionComponents(args, dynamicsApiUrl);

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

  private async getSolutionComponents(args: CommandArgs, dynamicsApiUrl: string): Promise<string> {
    const solutionId = await this.getSolutionId(args);
    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${solutionId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: SolutionComponent[] }>(requestOptions);

    return this.formatAsXml(response.value);
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

  private formatAsXml(solutionComponents: SolutionComponent[]): string {
    const result = solutionComponents.reduce(function (r, a) {
      const key = a.msdyn_componentlogicalname.slice(-1) === 'y' ?
        a.msdyn_componentlogicalname.substring(0, a.msdyn_componentlogicalname.length - 1) + 'ies' :
        a.msdyn_componentlogicalname + 's';
      r[key] = r[key] || [];

      r[key].push({ [a.msdyn_componentlogicalname]: a.msdyn_name });
      return r;
    }, Object.create(null));

    return `<importexportxml>${this.objectToXml(result)}</importexportxml>`;
  }

  private objectToXml(obj: any): string {
    let xml = '';
    for (const prop in obj) {
      xml += "<" + prop + ">";
      if (obj[prop] instanceof Array) {
        for (const array in obj[prop]) {
          xml += this.objectToXml(new Object(obj[prop][array]));
        }
      }
      else {
        xml += obj[prop];
      }
      xml += "</" + prop + ">";
    }
    xml = xml.replace(/<\/?[0-9]{1,}>/g, '');
    return xml;
  }

}

module.exports = new PpSolutionPublishCommand();