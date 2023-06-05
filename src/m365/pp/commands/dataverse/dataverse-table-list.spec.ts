import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./dataverse-table-list');

describe(commands.DATAVERSE_TABLE_LIST, () => {
  //#region Mocked Responses
  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const dataverseResponse: any = {
    "value": {
      "@odata.context": "https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/$metadata#EntityDefinitions(MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex)",
      "value": [
        {
          "MetadataId": "27774349-6c36-44ab-8d5d-360df562cdd8",
          "IsCustomEntity": true,
          "IsManaged": true,
          "SchemaName": "aaduser",
          "IconVectorName": null,
          "LogicalName": "aaduser",
          "EntitySetName": "aadusers",
          "IsActivity": false,
          "DataProviderId": "54629ed7-0cd3-4c85-9b6c-ea5f8548a9aa",
          "IsRenameable": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "isrenameable"
          },
          "IsCustomizable": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "iscustomizable"
          },
          "CanCreateForms": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateforms"
          },
          "CanCreateViews": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateviews"
          },
          "CanCreateCharts": {
            "Value": false,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreatecharts"
          },
          "CanCreateAttributes": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateattributes"
          },
          "CanChangeTrackingBeEnabled": {
            "Value": false,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "canchangetrackingbeenabled"
          },
          "CanModifyAdditionalSettings": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canmodifyadditionalsettings"
          },
          "CanChangeHierarchicalRelationship": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canchangehierarchicalrelationship"
          },
          "CanEnableSyncToExternalSearchIndex": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canenablesynctoexternalsearchindex"
          }
        },
        {
          "MetadataId": "70816501-edb9-4740-a16c-6a5efbc05d84",
          "IsCustomEntity": false,
          "IsManaged": true,
          "SchemaName": "Account",
          "IconVectorName": null,
          "LogicalName": "account",
          "EntitySetName": "accounts",
          "IsActivity": false,
          "DataProviderId": null,
          "IsRenameable": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "isrenameable"
          },
          "IsCustomizable": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "iscustomizable"
          },
          "CanCreateForms": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateforms"
          },
          "CanCreateViews": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateviews"
          },
          "CanCreateCharts": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreatecharts"
          },
          "CanCreateAttributes": {
            "Value": true,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "cancreateattributes"
          },
          "CanChangeTrackingBeEnabled": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canchangetrackingbeenabled"
          },
          "CanModifyAdditionalSettings": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canmodifyadditionalsettings"
          },
          "CanChangeHierarchicalRelationship": {
            "Value": false,
            "CanBeChanged": false,
            "ManagedPropertyLogicalName": "canchangehierarchicalrelationship"
          },
          "CanEnableSyncToExternalSearchIndex": {
            "Value": true,
            "CanBeChanged": true,
            "ManagedPropertyLogicalName": "canenablesynctoexternalsearchindex"
          }
        }
      ]
    }
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DATAVERSE_TABLE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['SchemaName', 'EntitySetName', 'LogicalName', 'IsManaged']);
  });

  it('retrieves data from dataverse', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&$filter=(IsIntersect eq false and IsLogicalEntity eq false and%0APrimaryNameAttribute ne null and PrimaryNameAttribute ne %27%27 and ObjectTypeCode gt 0 and%0AObjectTypeCode ne 4712 and ObjectTypeCode ne 4724 and ObjectTypeCode ne 9933 and ObjectTypeCode ne 9934 and%0AObjectTypeCode ne 9935 and ObjectTypeCode ne 9947 and ObjectTypeCode ne 9945 and ObjectTypeCode ne 9944 and%0AObjectTypeCode ne 9942 and ObjectTypeCode ne 9951 and ObjectTypeCode ne 2016 and ObjectTypeCode ne 9949 and%0AObjectTypeCode ne 9866 and ObjectTypeCode ne 9867 and ObjectTypeCode ne 9868) and (IsCustomizable/Value eq true or IsCustomEntity eq true or IsManaged eq false or IsMappable/Value eq true or IsRenameable/Value eq true)&api-version=9.1`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return dataverseResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36' } });
    assert(loggerLogSpy.calledWith(dataverseResponse.value));
  });

  it('retrieves data from dataverse as admin', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&$filter=(IsIntersect eq false and IsLogicalEntity eq false and%0APrimaryNameAttribute ne null and PrimaryNameAttribute ne %27%27 and ObjectTypeCode gt 0 and%0AObjectTypeCode ne 4712 and ObjectTypeCode ne 4724 and ObjectTypeCode ne 9933 and ObjectTypeCode ne 9934 and%0AObjectTypeCode ne 9935 and ObjectTypeCode ne 9947 and ObjectTypeCode ne 9945 and ObjectTypeCode ne 9944 and%0AObjectTypeCode ne 9942 and ObjectTypeCode ne 9951 and ObjectTypeCode ne 2016 and ObjectTypeCode ne 9949 and%0AObjectTypeCode ne 9866 and ObjectTypeCode ne 9867 and ObjectTypeCode ne 9868) and (IsCustomizable/Value eq true or IsCustomEntity eq true or IsManaged eq false or IsMappable/Value eq true or IsRenameable/Value eq true)&api-version=9.1`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return dataverseResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', asAdmin: true } });
    assert(loggerLogSpy.calledWith(dataverseResponse.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      throw {
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      };
    });

    try {
      await command.action(logger, { options: { environment: '4be50206-9576-4237-8b17-38d8aadfaa36' } });
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert(ex, (new CommandError("Resource '' does not exist or one of its queried reference-property objects are not present")).message);
    }
  });
});
