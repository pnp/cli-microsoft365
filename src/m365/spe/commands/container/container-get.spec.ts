import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
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
  const containerId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ';
  const containerName = 'My Application Storage Container';
  const containerResponse = {
    id: containerId,
    displayName: containerName,
    description: 'Description of My Application Storage Container',
    containerTypeId: '91710488-5756-407f-9046-fbe5f0b4de73',
    status: 'active',
    createdDateTime: '2021-11-24T15:41:52.347Z',
    settings: {
      isOcrEnabled: false
    }
  };
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let schema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    schema = command.getSchemaToParse()!;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    loggerLogSpy.restore();
    loggerLogToStderrSpy.restore();
    sinonUtil.restore([request.get]);
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return containerResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: containerId } } as any);
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], containerResponse);
  });

  it('gets container by name', async () => {
    sinon.stub(request, 'get').onFirstCall().resolves({
      value: [containerResponse]
    }).onSecondCall().callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return containerResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: containerName } } as any);
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], containerResponse);
  });

  it('fails when container with specified name does not exist', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await assert.rejects(
      command.action(logger, { options: { name: containerName } } as any),
      new CommandError(`Container with name '${containerName}' not found.`)
    );
  });

  it('logs progress when resolving container id by name in verbose mode', async () => {
    sinon.stub(request, 'get').onFirstCall().resolves({
      value: [containerResponse]
    }).onSecondCall().callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return containerResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: containerName, verbose: true } } as any);
    assert(loggerLogToStderrSpy.calledWith(`Resolving container id from name '${containerName}'...`));
  });

  it('throws received error when resolving container id fails with unexpected error', async () => {
    const unexpectedError = new Error('Unexpected');
    sinon.stub(request, 'get').rejects(unexpectedError);

    try {
      await command.action(logger, { options: { name: containerName } } as any);
      assert.fail('Expected command to throw');
    }
    catch (err: any) {
      assert.strictEqual(err.message, unexpectedError.message);
    }
  });

  it('fails validation when neither id nor name is specified', () => {
    const result = schema.safeParse({});
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('Specify either id or name')));
  });

  it('fails validation when both id and name are specified', () => {
    const result = schema.safeParse({ id: containerId, name: containerName });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('Specify either id or name')));
  });

  it('passes validation when only id is specified', () => {
    const result = schema.safeParse({ id: containerId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when only name is specified', () => {
    const result = schema.safeParse({ name: containerName });
    assert.strictEqual(result.success, true);
  });
});
