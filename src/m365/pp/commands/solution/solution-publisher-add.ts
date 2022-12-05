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
    return 'Adds a specified publisher in a given environment';
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
        if (isNaN(args.options.choiceValuePrefix) || args.options.choiceValuePrefix < 10000 || args.options.choiceValuePrefix > 99999 || !Number.isInteger(args.options.choiceValuePrefix)) {
          return 'choiceValuePrefix should be an integer between 10000 and 99999.';
        }

        const nameRegEx = new RegExp(/^[a-zA-Z_][A-Za-z0-9_]+$/);
        if (!nameRegEx.test(args.options.name)) {
          return 'option name may only consist of alphanumeric characters and underscores. The first character cannot be a number.';
        }

        const prefixRegEx = new RegExp(/^(?!mscrm.*$)[a-zA-Z_][A-Za-z0-9_]{1,7}$/);
        if (!prefixRegEx.test(args.options.prefix)) {
          return `optiong prefix may only consist of alphanumeric characters and underscores. The first character cannot be a number and may not start with 'mscrm'. Option prefix must be between 2 and 8 characters long.`;
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