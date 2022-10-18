import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import { Publisher, Solution } from './Solution';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
<<<<<<< HEAD
  id?: string;
  name?: string;
=======
  name: string;
>>>>>>> 9881501e (solution-get)
  asAdmin: boolean;
}

class PpSolutionGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_GET;
  }

  public get description(): string {
    return 'Lists a specific solution in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['uniquename', 'version', 'publisher'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
<<<<<<< HEAD
    // this.#initValidators();
    this.#initOptionSets();
=======
>>>>>>> 9881501e (solution-get)
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
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
<<<<<<< HEAD
        option: '-i, --id'
      },
      {
        option: '-n, --name'
=======
        option: '-n, --name <name>'
>>>>>>> 9881501e (solution-get)
      },
      {
        option: '-a, --asAdmin'
      }
    );
  }

<<<<<<< HEAD

  // #initValidators(): void {
  //   this.validators.push(
  //     async (args: CommandArgs) => {
  //       if (args.options.id && args.options.name) {
  //         return 'Specify either Id or Name but not both';
  //       }

  //       if (!args.options.id && !args.options.name) {
  //         return 'Specify either Id or Name';
  //       }

  //       return true;
  //     }
  //   );
  // }

  #initOptionSets(): void {
    this.optionSets.push(
      ['id', 'name']
    );
  }

=======
>>>>>>> 9881501e (solution-get)
  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving a specific solutions for which the user is an admin...`);
    }

    try {
<<<<<<< HEAD
      const res: Solution = await this.getSolution(args.options);
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

  private async getSolution(options: Options) {
    const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(options.environment, options.asAdmin);
    if (options.id) {
      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.0/solutions(${options.id})?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`,
=======
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq \'${args.options.name}\'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`,
>>>>>>> 9881501e (solution-get)
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

<<<<<<< HEAD
      const r = await request.get<Solution>(requestOptions);
      return r;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq \'${options.name}\'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const r = await request.get<{ value: Solution[] }>(requestOptions);
    return r.value[0];
=======
      const res = await request.get<{ value: Solution[] }>(requestOptions);

      if (!args.options.output || args.options.output === 'json') {
        logger.log(res.value[0]);
      }
      else {
        //converted to text friendly output
        if (res.value.length > 0) {
          const i = res.value[0];
          logger.log({
            uniquename: i.uniquename,
            version: i.version,
            publisher: (i.publisherid as Publisher).friendlyname
          });
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
>>>>>>> 9881501e (solution-get)
  }
}

module.exports = new PpSolutionGetCommand();