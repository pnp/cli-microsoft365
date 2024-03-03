import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './connection-get.js';

describe(commands.CONNECTION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnection: ExternalConnectors.ExternalConnection =
  {
    "id": "contosohr",
    "name": "Contoso HR",
    "description": "Connection to index Contoso HR system",
    "state": "draft",
    "configuration": {
      "authorizedAppIds": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ]
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.CONNECTION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
      }
    }), new CommandError('An error has occurred'));
  });

  it('should get external connection information by id (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return externalConnection;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: 'contosohr'
      }
    });

    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, 'contosohr');
    assert.strictEqual(call.args[0].name, 'Contoso HR');
    assert.strictEqual(call.args[0].description, 'Connection to index Contoso HR system');
    assert.strictEqual(call.args[0].state, 'draft');
  });

  it('should get external connection information by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return {
          "value": [
            externalConnection
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, 'contosohr');
    assert.strictEqual(call.args[0].name, 'Contoso HR');
    assert.strictEqual(call.args[0].description, 'Connection to index Contoso HR system');
    assert.strictEqual(call.args[0].state, 'draft');
  });

  it('fails retrieving external connection not found by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return {
          "value": []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    }), new CommandError(`External connection with name 'Contoso HR' not found`));
  });
});
