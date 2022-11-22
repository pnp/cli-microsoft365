import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environment: string;
  name: string;
  displayName: string;
  prefix: string;
  choiceValuePrefix: number;
  asAdmin?: boolean;
}

class PpSolutionPublisherAddCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_ADD;
  }

  public get description(): string {
    return 'Adds a specified publisher in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['publisherid', 'uniquename', 'friendlyname'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
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
        option: '-n, --name <name>'
      },
      {
        option: '--displayName <displayName>'
      },
      {
        option: '--prefix <prefix>'
      },
      {
        option: '--choiceValuePrefix <choiceValuePrefix>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.choiceValuePrefix && isNaN(args.options.choiceValuePrefix)) {
          return `choiceValuePrefix is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding a new publisher '${args.options.name}'...`);
    }
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.0/publishers`,
        headers: {
          'content-type': 'application/json;odata=verbose',
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          uniquename: args.options.name,
          friendlyname: args.options.displayName,
          customizationprefix: args.options.prefix,
          customizationoptionvalueprefix: args.options.choiceValuePrefix
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }


  }
}

module.exports = new PpSolutionPublisherAddCommand();