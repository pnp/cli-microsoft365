import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './recyclebinitem-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.OWNER_LIST, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';

  const deletedFlows = [
    {
      name: '26a9a283-af42-4c09-aa3e-60c3cc166b90',
      id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/26a9a283-af42-4c09-aa3e-60c3cc166b90',
      type: 'Microsoft.ProcessSimple/environments/flows',
      properties: {
        apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
        displayName: 'Invoicing flow',
        state: 'Deleted',
        createdTime: '2024-08-05T23:13:54Z',
        lastModifiedTime: '2024-08-05T23:14:00Z',
        flowSuspensionReason: 'None',
        environment: {
          name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
          type: 'Microsoft.ProcessSimple/environments',
          id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
        },
        definitionSummary: {
          triggers: [],
          actions: []
        },
        creator: {
          tenantId: 'a16e76a1-837f-4bf9-82dc-78874d18e434',
          objectId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
          userId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
          userType: 'ActiveDirectory'
        },
        flowFailureAlertSubscribed: false,
        isManaged: false,
        machineDescriptionData: {},
        flowOpenAiData: {
          isConsequential: false,
          isConsequentialFlagOverwritten: false
        }
      }
    },
    {
      name: '53768068-1dd5-4cc4-a26b-034bad10bfed',
      id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/53768068-1dd5-4cc4-a26b-034bad10bfed',
      type: 'Microsoft.ProcessSimple/environments/flows',
      properties: {
        apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
        displayName: 'Invoicing flow 2',
        state: 'Deleted',
        createdTime: '2024-08-05T23:13:54Z',
        lastModifiedTime: '2024-08-05T23:14:00Z',
        flowSuspensionReason: 'None',
        environment: {
          name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
          type: 'Microsoft.ProcessSimple/environments',
          id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
        },
        definitionSummary: {
          triggers: [],
          actions: []
        },
        creator: {
          tenantId: 'a16e76a1-837f-4bf9-82dc-78874d18e434',
          objectId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
          userId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
          userType: 'ActiveDirectory'
        },
        flowFailureAlertSubscribed: false,
        isManaged: false,
        machineDescriptionData: {},
        flowOpenAiData: {
          isConsequential: false,
          isConsequentialFlagOverwritten: false
        }
      }
    }
  ];

  const flowResponse = {
    value: [
      {
        name: '7bb4a726-2e02-4b88-ad34-8510bcbbcfa0',
        id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/7bb4a726-2e02-4b88-ad34-8510bcbbcfa0',
        type: 'Microsoft.ProcessSimple/environments/flows',
        properties: {
          apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
          displayName: 'Create a Planner task when a channel post starts with TODO',
          state: 'Started',
          createdTime: '2024-03-22T15:09:07Z',
          lastModifiedTime: '2024-03-22T15:09:07Z',
          flowSuspensionReason: 'None',
          templateName: '2d30c27107de4d0786be7a2b4574ae70',
          environment: {
            name: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5',
            type: 'Microsoft.ProcessSimple/environments',
            id: '/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5'
          },
          definitionSummary: {
            triggers: [],
            actions: []
          },
          creator: {
            tenantId: 'a16e76a1-837f-4bf9-82dc-78874d18e434',
            objectId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
            userId: 'bd51c64d-c262-4184-ba3f-5361ea553820',
            userType: 'ActiveDirectory'
          },
          provisioningMethod: 'FromTemplate',
          flowFailureAlertSubscribed: false,
          isManaged: false,
          machineDescriptionData: {},
          flowOpenAiData: {
            isConsequential: false,
            isConsequentialFlagOverwritten: false
          }
        }
      },
      ...deletedFlows
    ]
  };

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName']);
  });

  it('outputs exactly one result when retrieving deleted flows with output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName } });
    assert(loggerLogSpy.calledOnce);
  });

  it('outputs exactly one result when retrieving deleted flows with output text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, output: 'text' } });
    assert(loggerLogSpy.calledOnce);
  });

  it('correctly retrieves deleted flows with output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], deletedFlows);
  });

  it('correctly retrieves deleted flows with output text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`) {
        return flowResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, output: 'text' } });
    const textResponse = deletedFlows.map(flow => ({ ...flow, displayName: flow.properties.displayName }));

    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], textResponse);
  });

  it('throws error when no environment found', async () => {
    const error = {
      'error': {
        'code': 'EnvironmentAccessDenied',
        'message': `Access to the environment '${environmentName}' is denied.`
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName } }),
      new CommandError(error.error.message));
  });
});