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
import command from './container-recyclebinitem-remove.js';
import { spe } from '../../../../utils/spe.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTAINER_RECYCLEBINITEM_REMOVE, () => {
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';
  const containerId = 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND';
  const containerName = 'Container name';

  const deletedContainersResponse = [
    {
      id: 'b!dmyqSLRGPke6nU-Yi7sSsLAUFJYzrkJGh53mNOfkEfvkMQQfuvPfRLuZpE-wBQ6n',
      displayName: 'Deleted Container 1'
    },
    {
      id: containerId,
      displayName: containerName
    }
  ];

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_RECYCLEBINITEM_REMOVE);
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

  it('prompts before permanently removing the deleted container', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { id: containerTypeId } });
    assert(confirmationStub.calledOnce);
  });

  it('aborts permanently removing the deleted container when prompt is not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const deleteStub = sinon.stub(request, 'delete').resolves({});

    await command.action(logger, { options: { id: containerTypeId } });
    assert(deleteStub.notCalled);
  });

  it('permanently removes a deleted container by id', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('permanently removes a deleted container by name', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: deletedContainersResponse
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId } });
    assert(deleteStub.calledOnce);
    assert(confirmationStub.calledOnce);
  });

  it('permanently removes a deleted container by name and container type name', async () => {
    sinon.stub(spe, 'getContainerTypeIdByName').callsFake(async (name) => {
      if (name === containerTypeName) {
        return containerTypeId;
      }

      throw `Container type with name '${name}' not found.`;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: deletedContainersResponse
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeName: containerTypeName, verbose: true, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('correctly throws error when container name does not exist', async () => {
    sinon.stub(request, 'get').resolves({
      value: deletedContainersResponse
    });

    await assert.rejects(command.action(logger, { options: { name: 'Non-existing container', containerTypeId: containerTypeId, force: true } }),
      new CommandError(`The specified container 'Non-existing container' does not exist.`));
  });

  it('correctly handles multiple containers with the same name', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}`) {
        return;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: [
            ...deletedContainersResponse,
            {
              id: 'b!anotherContainerId',
              displayName: containerName
            }
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(deletedContainersResponse.find(c => c.id === containerId)!);
    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId, force: true } });
    assert(stubMultiResults.calledOnce);
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'delete').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerId, force: true } }),
      new CommandError(errorMessage));
  });
});