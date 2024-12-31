import assert from "assert";
import commands from "../../commands.js";
import command from './pipeline-list.js';
import sinon from 'sinon';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { pid } from '../../../../utils/pid.js';
import { session } from "../../../../utils/session.js";
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandError } from "../../../../Command.js";




describe(commands.PIPELINE_LIST, () => {
  const environmentName = 'environmentName';
  const mockPipelineListResponse: any = [
    {
      name: 'pipeline1',
      deploymentpipelineid: 'deploymentpipelineid1',
      ownerid: 'owner1',
      statuscode: 'statuscode1'
    }
  ];
  const mockEnvironmentResponse = {
    "id": `/providers/Microsoft.BusinessAppPlatform/environments/Default-Environment`,
    "type": "Microsoft.BusinessAppPlatform/environments",
    "location": "unitedstates",
    "name": "Default-Environment",
    "properties": {
      "displayName": "contoso (default)",
      "isDefault": true,
      linkedEnvironmentMetadata: {
        instanceApiUrl: 'https://contoso.crm.dynamics.com'
      }
    }
  };

  let log: string[];
  let logger: Logger;

  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(commands.PIPELINE_LIST, command.name);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'deploymentpipelineid', 'ownerid', 'statuscode']);
  });

  it('should send a request to get environment details', async () => {
    const getEnvironmentStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments/${environmentName}`) > -1) {
        return mockEnvironmentResponse;
      }
      throw 'Invalid request';
    });
    await command['getEnvironmentDetails'](environmentName, false);

    assert(getEnvironmentStub.calledWith(sinon.match({
      url: sinon.match(`https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environmentName}`),
      headers: sinon.match({
        accept: 'application/json;odata.metadata=none'
      }),
      responseType: 'json'
    })));
  });

  it('should send a request to get environment details (as admin)', async () => {
    const getEnvironmentStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentName}`) > -1) {
        return mockEnvironmentResponse;
      }
      throw 'Invalid request';
    });
    await command['getEnvironmentDetails'](environmentName, true);

    assert(getEnvironmentStub.calledWith(sinon.match({
      url: sinon.match(`https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${environmentName}`),
      headers: sinon.match({
        accept: 'application/json;odata.metadata=none'
      }),
      responseType: 'json'
    })));
  });

  it('returns correct environment details', async () => {
    sinon.stub(request, 'get').resolves(mockEnvironmentResponse);
    const result = await command['getEnvironmentDetails'](environmentName, false);
    assert.strictEqual(result.properties.linkedEnvironmentMetadata.instanceApiUrl, 'https://contoso.crm.dynamics.com');
  });

  it('retrieves pipelines in the specified Power Platform environment', async () => {
    const getEnvironmentStub = await sinon.stub(command as any, 'getEnvironmentDetails').callsFake(() => Promise.resolve(mockEnvironmentResponse));
    const getPipelineStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/api/data/v9.0/deploymentpipelines') > -1) {
        return {
          value: [
            {
              name: 'pipeline1',
              deploymentpipelineid: 'deploymentpipelineid1',
              '_ownerid_value': 'owner1',
              statuscode: 'statuscode1'
            }
          ]
        };
      }
      throw new Error('Invalid request');
    });

    await command.action(logger, {
      options: {
        environmentName: environmentName,
        asAdmin: false
      }
    });

    assert(getEnvironmentStub.called);
    assert(getPipelineStub.called);

    assert(loggerLogSpy.calledWith(sinon.match(mockPipelineListResponse)));
  });

  it('correctly handles error when retrieving environment details or pipelines', async () => {
    const errorMessage = 'An error has occurred';
    sinon.stub(request, 'get').callsFake(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: environmentName,
        asAdmin: false
      }
    }), new CommandError(errorMessage));
  });

});



