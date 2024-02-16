import { Group } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import fs from 'fs';
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
import command from './m365group-set.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.M365GROUP_SET, () => {
  const fsStats: fs.Stats = {
    isDirectory: () => false,
    isFile: () => false,
    isBlockDevice: () => false,
    isCharacterDevice: () => false,
    isSymbolicLink: () => false,
    isFIFO: () => false,
    isSocket: () => false,
    dev: 0,
    ino: 0,
    mode: 0,
    nlink: 0,
    uid: 0,
    gid: 0,
    rdev: 0,
    size: 0,
    blksize: 0,
    blocks: 0,
    atimeMs: 0,
    mtimeMs: 0,
    ctimeMs: 0,
    birthtimeMs: 0,
    atime: new Date(),
    mtime: new Date(),
    ctime: new Date(),
    birthtime: new Date()
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  const groupId = 'f3db5c2b-068f-480d-985b-ec78b9fa0e76';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).pollingInterval = 0;
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.put,
      request.patch,
      request.get,
      fs.readFileSync,
      fs.existsSync,
      fs.lstatSync,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_SET]);
  });

  it('updates Microsoft 365 Group display name', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.data) === JSON.stringify(<Group>{
          displayName: 'My group'
        })) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates Microsoft 365 Group description (debug)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.data) === JSON.stringify(<Group>{
          description: 'My group'
        })) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', description: 'My group' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('clears Microsoft 365 Group description when empty string is passed', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, description: '' } });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(patchStub.firstCall.args[0].data)), { description: null });
  });

  it('updates Microsoft 365 Group to public', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.data) === JSON.stringify(<Group>{
          visibility: 'Public'
        })) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: false } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates Microsoft 365 Group to private', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        if (JSON.stringify(opts.data) === JSON.stringify(<Group>{
          visibility: 'Private'
        })) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates Microsoft 365 Group logo with a png image', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/png') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'logo.png' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates Microsoft 365 Group logo with a jpg image (debug)', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/jpeg') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.jpg' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('updates Microsoft 365 Group logo with a gif image', async () => {
    sinon.stub(fs, 'readFileSync').returns('abc');
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/gif') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.gif' } });
    assert(loggerLogSpy.notCalled);
  });

  it('handles failure when updating Microsoft 365 Group logo and succeeds after 10 attempts', async () => {
    let amountOfCalls = 1;
    sinon.stub(fs, 'readFileSync').returns('abc');
    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' && amountOfCalls < 10) {
        amountOfCalls++;
        throw 'Invalid request';
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.png' } });
    assert.strictEqual(putStub.callCount, 10);
  });

  it('handles failure when updating Microsoft 365 Group logo', async () => {
    const error = {
      error: {
        message: 'An error has occurred'
      }
    };
    sinon.stub(fs, 'readFileSync').returns('abc');
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.png' } } as any),
      new CommandError('An error has occurred'));
  });

  it('handles failure when updating Microsoft 365 Group logo (debug)', async () => {
    const error = {
      error: {
        message: 'An error has occurred'
      }
    };
    sinon.stub(fs, 'readFileSync').returns('abc');
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', logoPath: 'logo.png' } } as any),
      new CommandError('An error has occurred'));
  });

  it('adds owner to Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', owners: 'user@contoso.onmicrosoft.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds owners to Microsoft 365 Group (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com' or userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            },
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('adds member to Microsoft 365 Group', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', members: 'user@contoso.onmicrosoft.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds members to Microsoft 365 Group (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com' or userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('sets option allowExternalSenders when using delegated permissions', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, allowExternalSenders: true } });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(patchStub.firstCall.args[0].data)), { allowExternalSenders: true });
  });

  it('sets option autoSubscribeNewMembers when using delegated permissions', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, autoSubscribeNewMembers: true } });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(patchStub.firstCall.args[0].data)), { autoSubscribeNewMembers: true });
  });

  it('sets option hideFromAddressLists and autoSubscribeNewMembers when using delegated permissions', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, autoSubscribeNewMembers: true, hideFromAddressLists: false } });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(patchStub.firstCall.args[0].data)), { autoSubscribeNewMembers: true, hideFromAddressLists: false });
  });

  it('sets option hideFromOutlookClients correctly', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, hideFromOutlookClients: true } });
    assert.deepStrictEqual(JSON.parse(JSON.stringify(patchStub.firstCall.args[0].data)), { hideFromOutlookClients: true });
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } } as any),
      new CommandError('An error has occurred'));
  });

  it('throws error when the group is not a unified group', async () => {
    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId, displayName: 'Updated title' } }),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });

  it('throws error when we are trying to update allowExternalSenders and we are using application only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').resolves(true);

    await assert.rejects(command.action(logger, { options: { id: groupId, allowExternalSenders: true } }),
      new CommandError(`Option 'allowExternalSenders' and 'autoSubscribeNewMembers' can only be used when using delegated permissions.`));
  });


  it('throws error when we are trying to update autoSubscribeNewMembers and we are using application only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').resolves(true);

    await assert.rejects(command.action(logger, { options: { id: groupId, autoSubscribeNewMembers: true } }),
      new CommandError(`Option 'allowExternalSenders' and 'autoSubscribeNewMembers' can only be used when using delegated permissions.`));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid', description: 'My awesome group' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID and displayName specified', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'My group' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID and description specified', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', description: 'My awesome group' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if no property to update is specified', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of the owners is invalid', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the owner is valid', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated with an additional space', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', owners: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if one of the members is invalid', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the member is valid', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated with an additional space', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', members: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if isPrivate is true', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if isPrivate is false', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if logoPath points to a non-existent file', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if logoPath points to a folder', async () => {
    const stats = { ...fsStats, isDirectory: () => true };
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'folder' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if logoPath points to an existing file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', logoPath: 'folder' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options are being set', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').returns(false);
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', displayName: 'Title', description: 'Description', logoPath: 'logo.png', owners: 'john@contoso.com', members: 'doe@contoso.com', isPrivate: false, allowExternalSenders: false, autoSubscribeNewMembers: false, hideFromAddressLists: false, hideFromOutlookClients: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
