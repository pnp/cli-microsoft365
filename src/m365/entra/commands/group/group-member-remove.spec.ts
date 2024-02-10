import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './group-member-remove.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { CommandError } from '../../../../Command.js';

describe(commands.GROUP_MEMBER_REMOVE, () => {
  const groupId = '630dfae3-6904-4154-acc2-812e11205351';
  const upns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a'];
  const groupNames = ['HR', 'Marketing', 'IT'];
  const groupIds = ['f64dc7f7-1a3e-4ba6-b4ee-491b282a3f84', '2e8641bb-9e9d-4da1-be52-d2a8394d3a85', '187a95ce-3e88-4051-87b9-ce19093975bf'];

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      entraGroup.getGroupIdByDisplayName,
      entraUser.getUserIdsByUpns,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'foo', userIds: userIds[0] } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userIds contains an invalid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userIds: `${userIds[0]},foo` } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userNames contains an invalid UPN', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userNames: `${upns[0]},foo` } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if subgroupIds contains an invalid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId, subgroupIds: `${groupIds[0]},foo`, role: 'Member' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role is not a valid role', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userIds: userIds.join(','), role: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if subgroupIds is specified without role option', async () => {
    const actual = await command.validate({ options: { groupId: groupId, subgroupIds: groupIds[0] } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if subgroupNames is specified with owner role', async () => {
    const actual = await command.validate({ options: { groupId: groupId, subgroupIds: groupIds[0], role: 'Owner' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with ids', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userIds: userIds.join(',') } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with ids with leading spaces', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userIds: userIds.map(i => ' ' + i).join(','), role: 'Member' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with names', async () => {
    const actual = await command.validate({ options: { groupName: 'IT department', userNames: upns.join(',') } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with names with trailing spaces', async () => {
    const actual = await command.validate({ options: { groupName: 'IT department', userNames: upns.map(u => u + ' ').join(','), role: 'Owner' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified users when confirm option not passed', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { groupName: 'IT department', subgroupNames: groupNames.join(','), role: 'Member' } });

    assert(confirmationStub.calledOnce);
  });

  it('aborts removing users when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const postSpy = sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { groupId: groupId, userIds: userIds.join(',') } });
    assert(postSpy.notCalled);
  });

  it('successfully removes owners and members from the group with ids after confirming prompt', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(20).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, userIds: userIds.join(','), verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 20 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/${index >= 10 ? 'members' : 'owners'}/${userIds[index % 10]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes owners and members from the group with ids with trailing spaces', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(20).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, userIds: userIds.map(i => i + ' ').join(','), force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 20 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/${index >= 10 ? 'members' : 'owners'}/${userIds[index % 10]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes owners and members from the group by using names after confirming', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(20).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Contoso', userNames: upns.join(','), verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 20 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/${index >= 10 ? 'members' : 'owners'}/${userIds[index % 10]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes owners and members from the group by using names with leading spaces', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(20).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Contoso', userNames: upns.map(u => + ' ' + u).join(','), force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 20 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/${index >= 10 ? 'members' : 'owners'}/${userIds[index % 10]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes owners from the group with ids', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(10).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, userIds: userIds.join(','), role: 'Owner', force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 10 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/owners/${userIds[index]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes members from the group by using names', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(10).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Contoso', userNames: upns.join(','), role: 'Member', force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 10 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/members/${userIds[index]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes subgroups from the group by using IDs', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(3).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, subgroupIds: groupIds.join(','), role: 'Member', force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 3 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/members/${groupIds[index]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('successfully removes subgroups from the group by using group names', async () => {
    const entraGroupStub = sinon.stub(entraGroup, 'getGroupIdByDisplayName').callsFake(async () => {
      return groupIds[entraGroupStub.callCount - 1];
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(3).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, subgroupNames: groupNames.join(','), role: 'Member', force: true, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, Array.from({ length: 3 }, (_, index) => ({
      id: index + 1,
      method: 'DELETE',
      url: `/groups/${groupId}/members/${groupIds[index]}/$ref`,
      headers: { 'content-type': 'application/json;odata.metadata=none' }
    })));
  });

  it('handles API errors correctly', async () => {
    const errorMessage = `Resource '${groupId}' does not exist or one of its queried reference-property objects are not present.`;

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array.from({ length: 20 }, (_, index) => {
            if (index < 10) {
              return {
                status: 204,
                body: {}
              };
            }
            return {
              status: 404,
              body: {
                error: {
                  code: 'Request_ResourceNotFound',
                  message: errorMessage
                }
              }
            };
          })
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { groupId: groupId, userIds: userIds.join(','), force: true, verbose: true } }),
      new CommandError(errorMessage));
  });

  it('correctly suppresses not found requests', async () => {
    const errorMessage = `Resource '${groupId}' does not exist or one of its queried reference-property objects are not present.`;

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array.from({ length: 20 }, (_, index) => {
            if (index < 10) {
              return {
                status: 204,
                body: {}
              };
            }
            return {
              status: 404,
              body: {
                error: {
                  code: 'Request_ResourceNotFound',
                  message: errorMessage
                }
              }
            };
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, userIds: userIds.join(','), suppressNotFound: true, force: true, verbose: true } });
    assert(postStub.calledOnce);
  });
});