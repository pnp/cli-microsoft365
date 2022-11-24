import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id: string;
  tablename: string;
  asAdmin?: boolean;
  confirm?: boolean;
}

class PpDataverseTableRowRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.DATAVERSE_TABLE_ROW_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific row from a dataverse table in the specified Power Platform environment.';
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
        asAdmin: !!args.options.asAdmin,
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --tablename <tablename>'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '--confirm'
      }
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
      logger.logToStderr(`Removing row '${args.options.id}' from table '${args.options.tablename}'...`);
    }

    if (args.options.confirm) {
      await this.deleteTableRow(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove row '${args.options.id}' from table '${args.options.tablename}'?`
      });

      if (result.continue) {
        await this.deleteTableRow(args);
      }
    }
  }


  private async deleteTableRow(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.1/${args.options.tablename}(${args.options.id})`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpDataverseTableRowRemoveCommand();