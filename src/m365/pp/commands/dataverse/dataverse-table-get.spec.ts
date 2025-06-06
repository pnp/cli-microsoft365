import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './dataverse-table-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.DATAVERSE_TABLE_GET, () => {
  //#region Mocked Responses
  const validName = "aaduser";
  const validEnvironment = "4be50206-9576-4237-8b17-38d8aadfaa36";
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const tableResponse: any = {
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
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DATAVERSE_TABLE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves data for a specific dataverse table', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='${validName}')?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return tableResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment, name: validName } });
    assert(loggerLogSpy.calledWith(tableResponse));
  });

  it('retrieves data from dataverse as admin', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='${validName}')?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return tableResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment, name: validName, asAdmin: true } });
    assert(loggerLogSpy.calledWith(tableResponse));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='${validName}')?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&api-version=9.1`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
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
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment, name: validName } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
