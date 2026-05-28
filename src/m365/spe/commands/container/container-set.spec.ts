import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './container-set.js';
import { z } from 'zod';

describe(commands.CONTAINER_SET, () => {
  const containerId = 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z';

  const patchResponse = {
    id: containerId,
    displayName: 'Updated Name',
    description: 'Updated Description',
    containerTypeId: '91710488-5756-407f-9046-fbe5f0b4de73',
    status: 'active',
    createdDateTime: '2021-11-24T15:41:52.347Z',
    lockState: 'unlocked',
    settings: {
      isOcrEnabled: false,
      itemMajorVersionLimit: 50,
      isItemVersioningEnabled: true
    }
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let baseSchema: z.ZodTypeAny;
  let refinedSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;

    commandInfo = cli.getCommandInfo(command);
    baseSchema = commandInfo.command.getSchemaToParse()!;
    refinedSchema = commandInfo.command.getRefinedSchema!(baseSchema as any)!;
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
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if no update options are provided', () => {
    const actual = refinedSchema.safeParse({ id: containerId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if itemMajorVersionLimit is not a positive integer', () => {
    const actual = baseSchema.safeParse({ id: containerId, itemMajorVersionLimit: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if itemMajorVersionLimit is a decimal number', () => {
    const actual = baseSchema.safeParse({ id: containerId, itemMajorVersionLimit: 12.5 });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when newName is specified', () => {
    const actual = refinedSchema.safeParse({ id: containerId, newName: 'New Name' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when description is specified', () => {
    const actual = refinedSchema.safeParse({ id: containerId, description: 'New description' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when isOcrEnabled is specified', () => {
    const actual = refinedSchema.safeParse({ id: containerId, isOcrEnabled: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when isItemVersioningEnabled is specified', () => {
    const actual = refinedSchema.safeParse({ id: containerId, isItemVersioningEnabled: false });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when itemMajorVersionLimit is specified', () => {
    const actual = refinedSchema.safeParse({ id: containerId, itemMajorVersionLimit: 100 });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all options are specified', () => {
    const actual = refinedSchema.safeParse({
      id: containerId,
      newName: 'New Name',
      description: 'New description',
      isOcrEnabled: true,
      isItemVersioningEnabled: true,
      itemMajorVersionLimit: 50
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly updates the container display name', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, newName: 'Updated Name' } });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      displayName: 'Updated Name'
    });
  });

  it('correctly updates the container description', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, description: 'Updated Description' } });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      description: 'Updated Description'
    });
  });

  it('correctly disables OCR for a container', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, isOcrEnabled: false } });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      settings: {
        isOcrEnabled: false
      }
    });
  });

  it('correctly enables versioning and sets the major version limit', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, isItemVersioningEnabled: true, itemMajorVersionLimit: 100 } });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      settings: {
        isItemVersioningEnabled: true,
        itemMajorVersionLimit: 100
      }
    });
  });

  it('correctly updates name, description, and settings together', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        id: containerId,
        newName: 'Updated Name',
        description: 'Updated Description',
        isOcrEnabled: true,
        isItemVersioningEnabled: true,
        itemMajorVersionLimit: 50,
        verbose: true
      }
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      displayName: 'Updated Name',
      description: 'Updated Description',
      settings: {
        isOcrEnabled: true,
        isItemVersioningEnabled: true,
        itemMajorVersionLimit: 50
      }
    });
  });

  it('correctly logs the output', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return patchResponse;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, newName: 'Updated Name' } });

    assert(loggerLogSpy.calledOnceWith(patchResponse));
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        code: 'accessDenied',
        message: 'Access denied'
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerId, newName: 'Updated Name' } }),
      new CommandError('Access denied'));
  });
});
