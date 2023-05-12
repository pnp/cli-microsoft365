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
const command: Command = require('./externalconnection-get');

describe(commands.EXTERNALCONNECTION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnection: any =
  {
    "id": "contosohr",
    "name": "Contoso HR",
    "description": "Connection to index Contoso HR system",
    "state": "draft",
    "configuration": {
      "authorizedApps": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ],
      "authorizedAppIds": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ]
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    (command as any).items = [];
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
      }
    }), new CommandError('An error has occurred'));
  });

  it('should get external connection information for the Microsoft Search by id (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return Promise.resolve(externalConnection);
      }
      return Promise.reject('Invalid request');
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

  it('should get external connection information for the Microsoft Search by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return Promise.resolve({
          "value": [
            externalConnection
          ]
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    }), new CommandError(`External connection with name 'Contoso HR' not found`));
  });
});
