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

  const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

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
      accessToken.isAppOnlyAccessToken,
      entraGroup.getGroupIdByDisplayName
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

  it('updates Microsoft 365 Group display name while group is being retrieved by display name', async () => {
    const groupName = 'Project A';
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: groupName, newDisplayName: 'My group', verbose: true } });
    assert(patchStub.calledOnce);
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

  it('adds members to Microsoft 365 Group by IDs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/microsoft.graph.user?$select=id`) {
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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'DELETE') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', memberIds: '949b16c1-a032-453e-a8ae-89a52bfc1d8a', verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds members to Microsoft 365 Group by UPNs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76?$select=groupTypes`) {
        return {
          groupTypes: [
            'Unified'
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/microsoft.graph.user?$select=id') {
        return {
          "value": [
            { "id": "949b16c1-a032-453e-a8ae-89a52bfc1d8a", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'DELETE') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'GET') {
        return {
          responses: [
            {
              id: userIds[0],
              status: 200,
              body: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', memberUserNames: 'user@contoso.onmicrosoft.com', verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds owners to Microsoft 365 Group by IDs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/microsoft.graph.user?$select=id`) {
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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'DELETE') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', ownerIds: '3527dada-9368-4cdd-a958-5460f5658e0e', verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('adds owners to Microsoft 365 Group by UPNs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76?$select=groupTypes`) {
        return {
          groupTypes: [
            'Unified'
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/microsoft.graph.user?$select=id') {
        return {
          "value": [
            { "id": "949b16c1-a032-453e-a8ae-89a52bfc1d8a", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'DELETE') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'GET') {
        return {
          responses: [
            {
              id: userIds[0],
              status: 200,
              body: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'f3db5c2b-068f-480d-985b-ec78b9fa0e76', ownerUserNames: 'user@contoso.onmicrosoft.com', verbose: true } });
    assert(loggerLogSpy.notCalled);
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

  it('handles API error when adding users to a group', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });
    sinon.stub(request, 'patch').resolves();
    sinon.stub(request, 'post').callsFake(async () => {
      return {
        responses: [
          {
            id: 1,
            status: 204,
            body: {}
          },
          {
            id: 2,
            status: 400,
            body: {
              error: {
                message: `One or more added object references already exist for the following modified properties: 'members'.`
              }
            }
          }
        ]
      };
    });

    await assert.rejects(command.action(logger, { options: { id: groupId, ownerIds: userIds.join(',') } }),
      new CommandError(`One or more added object references already exist for the following modified properties: 'members'.`));
  });

  it('handles API error when removing users from a group', async () => {
    sinon.stub(request, 'get').resolves({ value: [{ id: '717f1683-00fa-488c-b68d-5d0051f6bcfa' }] });
    sinon.stub(request, 'patch').resolves();
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'DELETE') {
        return {
          responses: [
            {
              status: 500,
              body: {
                error: {
                  message: 'Service unavailable.'
                }
              }
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: groupId, ownerIds: userIds.join(',') } }),
      new CommandError('Service unavailable.'));
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

    await assert.rejects(command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', newDisplayName: 'My group' } } as any),
      new CommandError('An error has occurred'));
  });

  it('throws error when the group is not a unified group', async () => {
    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId, newDisplayName: 'Updated title' } }),
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
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', newDisplayName: 'My group' } }, commandInfo);
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

  it('fails validation if ownerIds contains invalid GUID', async () => {
    const ownerIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', ownerIds: ownerIds.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerUserNames contains invalid user principal name', async () => {
    const ownerUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', ownerUserNames: ownerUserNames.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberIds contains invalid GUID', async () => {
    const memberIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', memberIds: memberIds.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberUserNames contains invalid user principal name', async () => {
    const memberUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', memberUserNames: memberUserNames.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if isPrivate is true', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if isPrivate is false', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', isPrivate: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with ids', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', ownerIds: userIds.join(',') } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with user names', async () => {
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', memberUserNames: userUpns.join(',') } }, commandInfo);
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
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);
    const actual = await command.validate({ options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', newDisplayName: 'Title', description: 'Description', logoPath: 'logo.png', ownerIds: userIds.join(','), memberIds: userIds.join(','), isPrivate: false, allowExternalSenders: false, autoSubscribeNewMembers: false, hideFromAddressLists: false, hideFromOutlookClients: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
