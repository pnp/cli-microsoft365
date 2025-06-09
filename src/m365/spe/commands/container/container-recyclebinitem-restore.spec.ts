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
import command from './container-recyclebinitem-restore.js';
import { spe } from '../../../../utils/spe.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTAINER_RECYCLEBINITEM_RESTORE, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
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
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_RECYCLEBINITEM_RESTORE);
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

  it('correctly logs no result', async () => {
    sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { id: containerId } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly restores a container by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}/restore`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerId } });
    assert(postStub.calledOnce);
  });

  it('correctly restores a container by name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}/restore`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: deletedContainersResponse
        };
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId } });
    assert(postStub.calledOnce);
  });

  it('correctly restores a container by name and container type name', async () => {
    sinon.stub(spe, 'getContainerTypeIdByName').callsFake(async (spoAdminUrl, name) => {
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

      throw 'Invalid GET request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}/restore`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeName: containerTypeName, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('correctly throws error when container name does not exist', async () => {
    sinon.stub(request, 'get').resolves({
      value: deletedContainersResponse
    });

    await assert.rejects(command.action(logger, { options: { name: 'Non-existing container', containerTypeId: containerTypeId } }),
      new CommandError(`The specified container 'Non-existing container' does not exist.`));
  });

  it('correctly handles multiple containers with the same name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers/${containerId}/restore`) {
        return;
      }

      throw 'Invalid POST request: ' + opts.url;
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

      throw 'Invalid GET request: ' + opts.url;
    });

    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(deletedContainersResponse.find(c => c.id === containerId)!);
    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId } });
    assert(stubMultiResults.calledOnce);
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'post').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerId } }),
      new CommandError(errorMessage));
  });
});