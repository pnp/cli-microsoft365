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
import command, { options } from './container-permission-add.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { session } from '../../../../utils/session.js';
import { spe } from '../../../../utils/spe.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.CONTAINER_PERMISSION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';
  const containerId = 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND';
  const containerName = 'Container name';
  const userId = '12345678-90ab-cdef-1234-567890abcdef';
  const userName = 'john.doe@contoso.com';

  const containerPermissionResponse = {
    "id": "X2k6MCMuZnxtZW1iZXJzaGlwfGRlYnJhYkBuYWNoYW4zNjUub25taWNyb3NvZnQuY29t",
    "roles": [
      "owner"
    ],
    "grantedToV2": {
      "user": {
        "displayName": "Debra Berger",
        "email": "debra@contoso.onmicrosoft.com",
        "userPrincipalName": "debra@contoso.onmicrosoft.com"
      }
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    sinon.stub(spe, 'getContainerTypeIdByName').withArgs(containerTypeName).resolves(containerTypeId);
    sinon.stub(spe, 'getContainerIdByName').withArgs(containerTypeId, containerName).resolves(containerId);
    sinon.stub(entraUser, 'getUpnByUserId').withArgs(userId).resolves(userName);

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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_PERMISSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both containerId and containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, containerName: containerName, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither containerId nor containerName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeId options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, containerTypeId: containerTypeId, roles: 'reader', userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, containerTypeName: containerTypeName, roles: 'reader', userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerName and both containerTypeId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerName: containerName, containerTypeId: containerTypeId, containerTypeName: containerTypeName, roles: 'reader', userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if roles are not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both userId and userName are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'reader', userId: userId, userName: userName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither userId nor userName are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'reader' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'reader', userId: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'reader', userName: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if correct role is not passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerId: containerId, roles: 'foo', userId: userId });
    assert.strictEqual(actual.success, false);
  });

  it('correctly adds permissions to a user for a container by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`) {
        return containerPermissionResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ containerId: containerId, roles: 'reader', userId: userId, verbose: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roles: ['reader'],
      grantedToV2: {
        user: {
          userPrincipalName: userName
        }
      }
    });
  });

  it('correctly adds permissions to a user for a container by name and container type by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`) {
        return containerPermissionResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ containerName: containerName, containerTypeId: containerTypeId, roles: 'reader', userId: userId, verbose: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roles: ['reader'],
      grantedToV2: {
        user: {
          userPrincipalName: userName
        }
      }
    });
  });

  it('correctly adds permissions to a user by UPN for a container by name and container type by name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`) {
        return containerPermissionResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ containerName: containerName, containerTypeName: containerTypeName, roles: 'reader', userName: userName, verbose: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roles: ['reader'],
      grantedToV2: {
        user: {
          userPrincipalName: userName
        }
      }
    });
  });

  it('correctly adds multiple permissions to a user for a container by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`) {
        return containerPermissionResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ containerId: containerId, roles: 'reader,writer', userId: userId, verbose: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      roles: ['reader', 'writer'],
      grantedToV2: {
        user: {
          userPrincipalName: userName
        }
      }
    });
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'post').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ containerId: containerId, roles: 'reader', userId: userId }) }),
      new CommandError(errorMessage));
  });
});