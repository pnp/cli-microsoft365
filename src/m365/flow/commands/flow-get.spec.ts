import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './flow-get.js';
import { flowGetResponse } from './flow-get.mock.js';

describe(commands.GET, () => {
  const expandUrl = '&$expand=swagger,properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    assert.strictEqual(command.name, commands.GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'displayName', 'description', 'triggers', 'actions']);
  });

  it('retrieves information about the specified flow (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith(flowGetResponse));
  });

  it('retrieves information about the specified flow', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith(flowGetResponse));
  });

  it('retrieves information about the specified flow (with trigger type)', async () => {
    const flowGetResponseClone = { ...flowGetResponse } as any;
    flowGetResponseClone.properties.definitionSummary.triggers = [
      {
        "type": "Request",
        "kind": "Button",
        "metadata": {
          "operationMetadataId": "a35030f3-0e3c-4501-aeba-0c9fb1b1d674"
        }
      }
    ];

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowGetResponseClone;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith(flowGetResponseClone));
  });

  it('retrieves information about the specified flow as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } });
    assert(loggerLogSpy.calledWith(flowGetResponse));
  });

  it('returns all properties for JSON output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return flowGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', output: 'json' } });
    assert(loggerLogSpy.calledWith(flowGetResponse));
  });

  it('renders empty string for description, if no description in the Flow specified', async () => {
    const flowGetResponseClone = { ...flowGetResponse } as any;
    flowGetResponseClone.properties.definitionSummary.description = undefined;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01${expandUrl}`) {
        return flowGetResponseClone;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert(loggerLogSpy.calledWith(flowGetResponseClone));
  });

  it('correctly handles no environment found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } } as any),
      new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
  });

  it('correctly handles Flow not found', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "ConnectionAuthorizationFailed",
        "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12' } } as any),
      new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'.`));
  });

  it('correctly handles Flow not found (as admin)', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "FlowNotFound",
        "message": "Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12', asAdmin: true } } as any),
      new CommandError(`Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } } as any),
      new CommandError('An error has occurred'));
  });
});
