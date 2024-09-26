
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-community-user-add.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.ENGAGE_COMMUNITY_USER_ADD, () => {
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIzNjAyMDAxMTAwOSJ9';
  const communityDisplayName = 'All company';
  const entraGroupId = 'b6c35b51-ebca-445c-885a-63a67d24cb53';
  const userNames = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves(userIds);
    sinon.stub(vivaEngage, 'getEntraGroupIdByCommunityDisplayName').resolves(entraGroupId);
    sinon.stub(vivaEngage, 'getEntraGroupIdByCommunityId').resolves(entraGroupId);
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_USER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if entraGroupId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: 'invalid',
      role: 'Member',
      userNames: userNames.join(',')
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if ids contains invalid guids', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      ids: userIds.join(',') + ',invalid',
      role: 'Member'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userNames contains invalid user principal names', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      userNames: userNames.join(',') + ',invalid',
      role: 'Member'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName or entraGroupId are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if communityId, communityDisplayName and entraGroupId are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      communityDisplayName: communityDisplayName,
      entraGroupId: entraGroupId,
      ids: userIds.join(','),
      role: 'Member'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if incorrect role value is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      userNames: userNames.join(','),
      role: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if communityId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      userNames: userNames.join(','),
      role: 'Admin'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if entraGroupId is specified with a proper GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      entraGroupId: entraGroupId,
      userNames: userNames.join(','),
      role: 'Admin'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if communityDisplayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      communityDisplayName: communityDisplayName,
      userNames: userNames.join(','),
      role: 'Admin'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if role is specified with a proper value', () => {
    const actual = commandOptionsSchema.safeParse({
      communityId: communityId,
      userNames: userNames.join(','),
      role: 'Admin'
    });
    assert.strictEqual(actual.success, true);
    assert(loggerLogSpy.notCalled);
  });

  it('correctly adds users specified by id as owner', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { communityDisplayName: communityDisplayName, verbose: true, ids: userIds.join(','), role: 'Owner' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
      {
        id: 1,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'owners@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      },
      {
        id: 21,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'owners@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      }
    ]);
  });

  it('correctly adds users specified by ids as member', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { communityId: communityId, verbose: true, ids: userIds.join(','), role: 'Member' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
      {
        id: 1,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'members@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      },
      {
        id: 21,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'members@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      }
    ]);
  });

  it('correctly adds users specified by userNames as member', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { entraGroupId: entraGroupId, verbose: true, userNames: userNames.join(','), role: 'Member' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, [
      {
        id: 1,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'members@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      },
      {
        id: 21,
        method: 'PATCH',
        url: `/groups/${entraGroupId}`,
        headers: { 'content-type': 'application/json;odata.metadata=none' },
        body: {
          'members@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
        }
      }
    ]);
  });

  it('handles API error when adding users to a community', async () => {
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

    await assert.rejects(command.action(logger, { options: { entraGroupId: entraGroupId, ids: userIds.join(','), role: 'Member' } }),
      new CommandError(`One or more added object references already exist for the following modified properties: 'members'.`));
  });
});