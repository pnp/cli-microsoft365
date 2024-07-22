import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './group-set.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { CommandError } from '../../../../Command.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.GROUP_SET, () => {
  const groupId = '7167b488-1ffb-43f1-9547-35969469bada';
  const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];
  const addOwnersRequest = [
    {
      id: 1,
      method: 'PATCH',
      url: `/groups/${groupId}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'owners@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    },
    {
      id: 21,
      method: 'PATCH',
      url: `/groups/${groupId}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'owners@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    }
  ];
  const addMembersRequest = [
    {
      id: 1,
      method: 'PATCH',
      url: `/groups/${groupId}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'members@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    },
    {
      id: 21,
      method: 'PATCH',
      url: `/groups/${groupId}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'members@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    }
  ];
  const updateGroupRequest = {
    displayName: '365 group',
    description: 'Microsoft 365 group',
    mailNickName: 'Microsoft365Group',
    visibility: 'Public'
  };
  const updateGroupWithEmptyDescriptionRequest = {
    displayName: '365 group',
    description: null,
    mailNickName: 'Microsoft365Group',
    visibility: 'Public'
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').withArgs('Microsoft 365 Group').resolves(groupId);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      request.post,
      entraUser.getUserIdsByUpns,
      entraGroup.getGroupIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the length of newDisplayName is more than 256 characters', async () => {
    const displayName = 'lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum';
    const actual = await command.validate({ options: { id: groupId, newDisplayName: displayName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the length of mailNickname is more than 64 characters', async () => {
    const mailNickname = 'loremipsumloremipsumloremipsumloremipsumloremipsumloremipsumloremipsumlorem';
    const actual = await command.validate({ options: { displayName: 'Cli group', mailNickname: mailNickname } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if mailNickname is not valid', async () => {
    const mailNickname = 'lorem ipsum';
    const actual = await command.validate({ options: { displayName: 'Cli group', mailNickname: mailNickname } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerIds contains invalid GUID', async () => {
    const ownerIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerIds: ownerIds.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerUserNames contains invalid user principal name', async () => {
    const ownerUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerUserNames: ownerUserNames.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberIds contains invalid GUID', async () => {
    const memberIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', memberIds: memberIds.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberUserNames contains invalid user principal name', async () => {
    const memberUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', memberUserNames: memberUserNames.join(',') } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if visibility contains invalid value', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', visibility: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: groupId, newDisplayName: 'Cli group' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'foo', newDisplayName: 'Cli group' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with ids', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerIds: userIds.join(',') } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with user names', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', memberUserNames: userUpns.join(',') } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if no options to be updated are specified', async () => {
    const actual = await command.validate({ options: { id: groupId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('successfully updates group specified by id', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });
    const patchRequestStub = sinon.stub(request, 'patch').resolves();

    await command.action(logger, { options: { id: groupId, description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', newDisplayName: '365 group', verbose: true } });
    assert.deepStrictEqual(patchRequestStub.lastCall.args[0].data, updateGroupRequest);
  });

  it('successfully updates group specified by displayName', async () => {
    sinon.stub(request, 'get').resolves({ value: []});

    const patchRequestStub = sinon.stub(request, 'patch').resolves();

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: '', mailNickname: 'Microsoft365Group', visibility: 'Public', newDisplayName: '365 group' } });
    assert.deepStrictEqual(patchRequestStub.lastCall.args[0].data, updateGroupWithEmptyDescriptionRequest);
  });

  it('successfully updates group with owners specified by ids and removes current owners', async () => {
    sinon.stub(request, 'get').resolves({ value: [{ id: '717f1683-00fa-488c-b68d-5d0051f6bcfa' }] });
    sinon.stub(request, 'patch').resolves();

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
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
              status: 204,
              body: {}
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', ownerIds: userIds.join(',') } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data.requests, addOwnersRequest);
  });

  it('successfully updates group with members specified by user names and removes current members', async () => {
    sinon.stub(request, 'get').resolves({ value: [{ id: '717f1683-00fa-488c-b68d-5d0051f6bcfa' }] });
    sinon.stub(entraUser, 'getUserIdsByUpns').withArgs(userUpns).resolves(userIds);
    sinon.stub(request, 'patch').resolves();

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
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
              status: 204,
              body: {}
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Private', memberUserNames: userUpns.join(','), verbose: true } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data.requests, addMembersRequest);
  });

  it('handles API error when adding users to a group', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });
    sinon.stub(request, 'patch').resolves();
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch' &&
        opts.data.requests[0].method === 'PATCH') {
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
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', ownerIds: userIds.join(',') } }),
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

    await assert.rejects(command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', ownerIds: userIds.join(',') } }),
      new CommandError('Service unavailable.'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: groupId, description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public' } }),
      new CommandError('Invalid request'));
  });
});