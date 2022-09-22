import { Logger } from '../../../../../cli';
import GlobalOptions from '../../../../../GlobalOptions';
import request from '../../../../../request';
import DataverseCommand from '../../../../base/DataverseCommand';
import commands from '../../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin: boolean;
}

class PpDataverseTableListCommand extends DataverseCommand {
  public get name(): string {
    return commands.DATAVERSE_TABLE_LIST;
  }

  public get description(): string {
    return 'Lists dataverse tables in given environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['SchemaName', 'EntitySetName', 'IsManaged'];
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
        option: '-a, --asAdmin'
      }
    );
  }

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of gateways for which the user is an admin...`);
    }

    this.getDynamicsInstance(args.options.environment, args.options.asAdmin)
      .then((dynamicsApiUrl: string) => {
        const requestOptions: any = {
          url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions?%24select=MetadataId%2CIsCustomEntity%2CIsManaged%2CSchemaName%2CIconVectorName%2CLogicalName%2CEntitySetName%2CIsActivity%2CDataProviderId%2CIsRenameable%2CIsCustomizable%2CCanCreateForms%2CCanCreateViews%2CCanCreateCharts%2CCanCreateAttributes%2CCanChangeTrackingBeEnabled%2CCanModifyAdditionalSettings%2CCanChangeHierarchicalRelationship%2CCanEnableSyncToExternalSearchIndex&%24filter=(IsIntersect%20eq%20false%20and%20IsLogicalEntity%20eq%20false%20and%0APrimaryNameAttribute%20ne%20null%20and%20PrimaryNameAttribute%20ne%20%27%27%20and%20ObjectTypeCode%20gt%200%20and%0AObjectTypeCode%20ne%204712%20and%20ObjectTypeCode%20ne%204724%20and%20ObjectTypeCode%20ne%209933%20and%20ObjectTypeCode%20ne%209934%20and%0AObjectTypeCode%20ne%209935%20and%20ObjectTypeCode%20ne%209947%20and%20ObjectTypeCode%20ne%209945%20and%20ObjectTypeCode%20ne%209944%20and%0AObjectTypeCode%20ne%209942%20and%20ObjectTypeCode%20ne%209951%20and%20ObjectTypeCode%20ne%202016%20and%20ObjectTypeCode%20ne%209949%20and%0AObjectTypeCode%20ne%209866%20and%20ObjectTypeCode%20ne%209867%20and%20ObjectTypeCode%20ne%209868)%20and%20(IsCustomizable%2FValue%20eq%20true%20or%20IsCustomEntity%20eq%20true%20or%20IsManaged%20eq%20false%20or%20IsMappable%2FValue%20eq%20true%20or%20IsRenameable%2FValue%20eq%20true)&api-version=9.1`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        request
          .get<{ value: any[] }>(requestOptions)
          .then((res: { value: any[] }): void => {
            logger.log(res.value);
            cb();
          });
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }


}

module.exports = new PpDataverseTableListCommand();
