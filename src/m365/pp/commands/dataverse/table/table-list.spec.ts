import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../../appInsights';
import auth from '../../../../../Auth';
import { Logger } from '../../../../../cli';
import Command, { CommandError } from '../../../../../Command';

import request from '../../../../../request';
import { sinonUtil } from '../../../../../utils';
import commands from '../../../commands';
const command: Command = require('./table-list');

describe(commands.DATAVERSE_TABLE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.DATAVERSE_TABLE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['SchemaName', 'EntitySetName', 'IsManaged']);
  });

  it('Retrieves retrieves data from dataverse as admin', async () => {
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

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions?%24select=MetadataId%2CIsCustomEntity%2CIsManaged%2CSchemaName%2CIconVectorName%2CLogicalName%2CEntitySetName%2CIsActivity%2CDataProviderId%2CIsRenameable%2CIsCustomizable%2CCanCreateForms%2CCanCreateViews%2CCanCreateCharts%2CCanCreateAttributes%2CCanChangeTrackingBeEnabled%2CCanModifyAdditionalSettings%2CCanChangeHierarchicalRelationship%2CCanEnableSyncToExternalSearchIndex&%24filter=(IsIntersect%20eq%20false%20and%20IsLogicalEntity%20eq%20false%20and%0APrimaryNameAttribute%20ne%20null%20and%20PrimaryNameAttribute%20ne%20%27%27%20and%20ObjectTypeCode%20gt%200%20and%0AObjectTypeCode%20ne%204712%20and%20ObjectTypeCode%20ne%204724%20and%20ObjectTypeCode%20ne%209933%20and%20ObjectTypeCode%20ne%209934%20and%0AObjectTypeCode%20ne%209935%20and%20ObjectTypeCode%20ne%209947%20and%20ObjectTypeCode%20ne%209945%20and%20ObjectTypeCode%20ne%209944%20and%0AObjectTypeCode%20ne%209942%20and%20ObjectTypeCode%20ne%209951%20and%20ObjectTypeCode%20ne%202016%20and%20ObjectTypeCode%20ne%209949%20and%0AObjectTypeCode%20ne%209866%20and%20ObjectTypeCode%20ne%209867%20and%20ObjectTypeCode%20ne%209868)%20and%20(IsCustomizable%2FValue%20eq%20true%20or%20IsCustomEntity%20eq%20true%20or%20IsManaged%20eq%20false%20or%20IsMappable%2FValue%20eq%20true%20or%20IsRenameable%2FValue%20eq%20true)&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(dataverseResponse);
        }
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', asAdmin: true } });
    assert(loggerLogSpy.calledWith(dataverseResponse.value));
  });

  it('Retrieves retrieves data from dataverse', async () => {
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

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions?%24select=MetadataId%2CIsCustomEntity%2CIsManaged%2CSchemaName%2CIconVectorName%2CLogicalName%2CEntitySetName%2CIsActivity%2CDataProviderId%2CIsRenameable%2CIsCustomizable%2CCanCreateForms%2CCanCreateViews%2CCanCreateCharts%2CCanCreateAttributes%2CCanChangeTrackingBeEnabled%2CCanModifyAdditionalSettings%2CCanChangeHierarchicalRelationship%2CCanEnableSyncToExternalSearchIndex&%24filter=(IsIntersect%20eq%20false%20and%20IsLogicalEntity%20eq%20false%20and%0APrimaryNameAttribute%20ne%20null%20and%20PrimaryNameAttribute%20ne%20%27%27%20and%20ObjectTypeCode%20gt%200%20and%0AObjectTypeCode%20ne%204712%20and%20ObjectTypeCode%20ne%204724%20and%20ObjectTypeCode%20ne%209933%20and%20ObjectTypeCode%20ne%209934%20and%0AObjectTypeCode%20ne%209935%20and%20ObjectTypeCode%20ne%209947%20and%20ObjectTypeCode%20ne%209945%20and%20ObjectTypeCode%20ne%209944%20and%0AObjectTypeCode%20ne%209942%20and%20ObjectTypeCode%20ne%209951%20and%20ObjectTypeCode%20ne%202016%20and%20ObjectTypeCode%20ne%209949%20and%0AObjectTypeCode%20ne%209866%20and%20ObjectTypeCode%20ne%209867%20and%20ObjectTypeCode%20ne%209868)%20and%20(IsCustomizable%2FValue%20eq%20true%20or%20IsCustomEntity%20eq%20true%20or%20IsManaged%20eq%20false%20or%20IsMappable%2FValue%20eq%20true%20or%20IsRenameable%2FValue%20eq%20true)&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(dataverseResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36' } });
    assert(loggerLogSpy.calledWith(dataverseResponse.value));
  });

  it('correctly handles access denied to environment', async () => {
    const errorResponse: any = {
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    };

    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject(errorResponse);
    });

    await assert.rejects(command.action(logger, { options: { } } as any),
      new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
  });

  it('correctly handles non existing environments', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentNotFound",
          "message": "The environment 'nonexisting' could not be found in the tenant 'someid'."
        }
      });
    });


    await assert.rejects(command.action(logger, { options: { debug: false, environment: 'nonexisting' } } as any),
      new CommandError(`The environment 'nonexisting' could not be found in the tenant 'someid'.`));
  });

  it('correctly handles dataverse URI not found', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/noDynamics?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, environment: 'noDynamics' } } as any),
      new CommandError(`No Dynamics instance found for 'noDynamics'`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));

  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
