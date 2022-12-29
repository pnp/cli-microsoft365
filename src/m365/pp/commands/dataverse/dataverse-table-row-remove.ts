import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id: string;
  entitySetName?: string;
  tableName?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        entitySetName: typeof args.options.entitySetName !== 'undefined',
        tableName: typeof args.options.tableName !== 'undefined',
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
        option: '--entitySetName [entitySetName]'
      },
      {
        option: '--tableName [tableName]'
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

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['entitySetName', 'tableName'] }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'...`);
    }

    if (args.options.confirm) {
      await this.deleteTableRow(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'?`
      });

      if (result.continue) {
        await this.deleteTableRow(logger, args);
      }
    }
  }

  private async deleteTableRow(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
      if (this.verbose) {
        logger.logToStderr('Entity set name is: ' + entitySetName);
      }

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/${entitySetName}(${args.options.id})`,
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

  protected async getEntitySetName(dynamicsApiUrl: string, args: CommandArgs): Promise<string> {
    if (args.options.entitySetName) {
      return args.options.entitySetName;
    }

    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${args.options.tableName}')?$select=EntitySetName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ EntitySetName: string }>(requestOptions);

    return response.EntitySetName;
  }
}

module.exports = new PpDataverseTableRowRemoveCommand();