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
import command, { options } from './container-permission-remove.js';
import { formatting } from '../../../../utils/formatting.js';
import { session } from '../../../../utils/session.js';
import { spe } from '../../../../utils/spe.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.CONTAINER_PERMISSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean;

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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_PERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if permission id is not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both containerId and containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerName: containerName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither containerId nor containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeId options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerTypeId: containerTypeId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerId: containerId, containerTypeName: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerName and both containerTypeId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: permissionId, containerName: containerName, containerTypeId: containerTypeId, containerTypeName: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('correctly removes permissions for a container by id', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid DELETE request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId, verbose: true, force: true }) });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes permissions for a container by name and container type by id and prompts for confirmation', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid DELETE request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerName: containerName, containerTypeId: containerTypeId, verbose: true }) });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes permissions for a container by name and container type by name', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${permissionId}`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerName: containerName, containerTypeName: containerTypeName, verbose: true, force: true }) });
    assert(deleteStub.calledOnce);
  });

  it('prompts before removing permissions when confirm option not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId }) });

    assert(promptIssued);
  });

  it('aborts removing permissions when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId }) });
    assert(deleteSpy.notCalled);
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'delete').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: permissionId, containerId: containerId, force: true }) }),
      new CommandError(errorMessage));
  });
});