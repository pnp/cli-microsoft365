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

interface Options extends GlobalOptions {
  environment: string;
  name: string;
  asAdmin: boolean;
}

class PpDataverseTableGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.DATAVERSE_TABLE_GET;
  }

  public get description(): string {
    return 'Gets a dataverse table in a given environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['SchemaName', 'EntitySetName', 'LogicalName', 'IsManaged'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
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
        option: '-a, --asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving a table for which the user is an admin...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const requestOptions: AxiosRequestConfig = {
        url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${args.options.name}')?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&api-version=9.1`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpDataverseTableGetCommand();