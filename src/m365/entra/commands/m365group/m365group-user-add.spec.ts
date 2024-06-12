import assert from 'assert';
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
import teamsCommands from '../../../teams/commands.js';
import commands from '../../commands.js';
import command from './m365group-user-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import aadCommands from '../../aadCommands.js';

describe(commands.M365GROUP_USER_ADD, () => {
  const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';
  const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue,
      entraGroup.getGroupIdByDisplayName,
      entraUser.getUserIdsByUpns
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_USER_ADD);
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
    assert.deepStrictEqual(alias, [teamsCommands.USER_ADD, aadCommands.M365GROUP_USER_ADD]);
  });

  it('fails validation if the groupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the groupId nor teamId are provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        role: 'Member',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both groupId and teamId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ids contains an invalid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId, ids: `${userIds[0]},foo`, role: 'Member' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userNames contains an invalid UPN', async () => {
    const actual = await command.validate({ options: { groupId: groupId, userNames: `${userUpns[0]},foo`, role: 'Member' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid groupId, userName and no role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId, userName and Owner role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'owner'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId, userName and Member role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'member'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds owners to specified Microsoft 365 group', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    sinon.stub(entraUser, 'getUserIdsByUpns').withArgs(userUpns).resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Contoso', userNames: userUpns.join(','), role: 'Owner', verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
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
    ]);
  });

  it('correctly adds members to specified Microsoft Teams team', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: groupId, ids: userIds.join(','), role: 'Member', verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
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
    ]);
  });

  it('correctly adds users to specified Microsoft Teams team with name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamName: 'Contoso', ids: userIds.join(','), verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
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
    ]);
  });

  it('handles API error when adding users to a group', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
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

    await assert.rejects(command.action(logger, { options: { groupId: groupId, ids: userIds.join(',') } }),
      new CommandError(`One or more added object references already exist for the following modified properties: 'members'.`));
  });

  it('handles error when the group is not a unified group', async () => {
    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId, userName: 'anne.matthews@contoso.onmicrosoft.com' } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
