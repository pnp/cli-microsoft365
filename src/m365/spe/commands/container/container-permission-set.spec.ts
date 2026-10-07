import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './container-permission-set.js';
import { formatting } from '../../../../utils/formatting.js';
import { session } from '../../../../utils/session.js';
import { spe } from '../../../../utils/spe.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.CONTAINER_PERMISSION_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const permissionId = 'cmVhZGVyX2k6MCMuZnxtZW1iZXJzaGlwfHJvcnlicjExMUBvdXRsb29rLmNvbQ';
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';
  const containerId = 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND';
  const containerName = 'Container name';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    sinon.stub(spe, 'getContainerTypeIdByName').withArgs(containerTypeName).resolves(containerTypeId);
    sinon.stub(spe, 'getContainerIdByName').withArgs(containerTypeId, containerName).resolves(containerId);

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.CONTAINER_PERMISSION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if permission id is not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both containerId and containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerName: containerName, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither containerId nor containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeId options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerTypeId: containerTypeId, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerTypeName: containerTypeName, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerName and both containerTypeId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerName: containerName, containerTypeId: containerTypeId, containerTypeName: containerTypeName, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if roles are not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if correct role is not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, roles: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('correctly updates permissions for a container by id', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId, roles: 'reader', verbose: true }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      roles: ['reader']
    });
  });

  it('correctly updates permissions for a container by name and container type by id', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerName: containerName, containerTypeId: containerTypeId, roles: 'reader', verbose: true }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      roles: ['reader']
    });
  });

  it('correctly updates permissions for a container by name and container type by name', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerName: containerName, containerTypeName: containerTypeName, roles: 'reader', verbose: true }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      roles: ['reader']
    });
  });

  it('correctly updates multiple permissions for a container by id', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId, roles: 'reader,writer', verbose: true }) });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      roles: ['reader', 'writer']
    });
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'patch').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId, roles: 'reader' }) }),
      new CommandError(errorMessage));
  });
});