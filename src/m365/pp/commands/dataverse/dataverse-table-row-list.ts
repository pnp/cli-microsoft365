import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environmentName: string;
  entitySetName?: string;
  tableName?: string;
  asAdmin?: boolean;
}

class PpDataverseTableRowListCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.DATAVERSE_TABLE_ROW_LIST;
  }

  public get description(): string {
    return 'Lists table rows for the given Dataverse table';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        entitySetName: typeof args.options.entitySetName !== 'undefined',
        tableName: typeof args.options.tableName !== 'undefined',
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
        option: '--entitySetName [entitySetName]'
      },
      {
        option: '--tableName [tableName]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['entitySetName', 'tableName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving list of table rows');
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
      if (this.verbose) {
        await logger.logToStderr('Entity set name is: ' + entitySetName);
      }

      const response = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.0/${entitySetName}`);
      await logger.log(response);
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

export default new PpDataverseTableRowListCommand();