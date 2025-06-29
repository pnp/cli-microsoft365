import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command from './container-remove.js';
import { spe } from '../../../../utils/spe.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTAINER_REMOVE, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';
  const containerId = 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND';
  const containerName = 'Container name';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    sinon.stub(spe, 'getContainerTypeIdByName').withArgs(spoAdminUrl, containerTypeName).resolves(containerTypeId);
    sinon.stub(spe, 'getContainerIdByName').withArgs(containerTypeId, containerName).resolves(containerId);

    auth.connection.active = true;
    auth.connection.spoUrl = spoAdminUrl.replace('-admin.sharepoint.com', '.sharepoint.com');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId, name: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither id nor name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerType option is used with id option', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId, containerTypeId: containerTypeId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if name is passed without containerType information', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerTypeId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeId: 'invalid-guid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if name is passed with containerTypeName', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeName: containerTypeName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if name is passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeId: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the container', async () => {
    await command.action(logger, { options: { id: containerTypeId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('prompts before recycling the container', async () => {
    await command.action(logger, { options: { id: containerTypeId, recycle: true } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the container when prompt is not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves([]);
    const deleteStub = sinon.stub(request, 'delete').resolves({});

    await command.action(logger, { options: { name: containerTypeName, containerTypeId: containerTypeId } });
    assert(postStub.notCalled);
    assert(deleteStub.notCalled);
  });

  it('correctly recycles a container by id', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return;
      }

      throw 'Invalid DELETE request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, recycle: true, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes a container by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}/permanentDelete`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, force: true } });
    assert(postStub.calledOnce);
  });

  it('correctly recycles a container by name', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`) {
        return;
      }

      throw 'Invalid DELETE request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId, recycle: true, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes a container by name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}/permanentDelete`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeName: containerTypeName, verbose: true, force: true } });
    assert(postStub.calledOnce);
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'post').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerId, force: true } }),
      new CommandError(errorMessage));
  });
});