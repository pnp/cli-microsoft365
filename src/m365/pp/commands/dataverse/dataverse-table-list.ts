import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  asAdmin: boolean;
}

class PpDataverseTableListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.DATAVERSE_TABLE_LIST;
  }

  public get description(): string {
    return 'Lists dataverse tables in a given environment';
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
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of tables for which the user is an admin...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const endpoint = `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&$filter=(IsIntersect eq false and IsLogicalEntity eq false and%0APrimaryNameAttribute ne null and PrimaryNameAttribute ne %27%27 and ObjectTypeCode gt 0 and%0AObjectTypeCode ne 4712 and ObjectTypeCode ne 4724 and ObjectTypeCode ne 9933 and ObjectTypeCode ne 9934 and%0AObjectTypeCode ne 9935 and ObjectTypeCode ne 9947 and ObjectTypeCode ne 9945 and ObjectTypeCode ne 9944 and%0AObjectTypeCode ne 9942 and ObjectTypeCode ne 9951 and ObjectTypeCode ne 2016 and ObjectTypeCode ne 9949 and%0AObjectTypeCode ne 9866 and ObjectTypeCode ne 9867 and ObjectTypeCode ne 9868) and (IsCustomizable/Value eq true or IsCustomEntity eq true or IsManaged eq false or IsMappable/Value eq true or IsRenameable/Value eq true)&api-version=9.1`;

      const res = await odata.getAllItems<any>(endpoint);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpDataverseTableListCommand();