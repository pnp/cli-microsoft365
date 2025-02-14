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
import command from './container-get.js';

describe(commands.CONTAINER_GET, () => {
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
    assert.strictEqual(command.name, commands.CONTAINER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    const errorMessage = 'Bad request.';
    sinon.stub(request, 'get').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: 'invalid', verbose: true } } as any),
      new CommandError(errorMessage));
  });

  it('gets container by id', async () => {
    const containerId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ';
    const response = {
      id: containerId,
      displayName: "My Application Storage Container",
      description: "Description of My Application Storage Container",
      containerTypeId: "91710488-5756-407f-9046-fbe5f0b4de73",
      status: "active",
      createdDateTime: "2021-11-24T15:41:52.347Z",
      settings: {
        isOcrEnabled: false
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: containerId } } as any);
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], response);
  });
});