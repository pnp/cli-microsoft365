import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./channel-member-set');

describe(commands.CHANNEL_MEMBER_SET, () => {
  const memberResponse = {
    "id": "00000",
    "roles": [],
    "displayName": "User",
    "userId": "00000000-0000-0000-0000-000000000000",
    "email": "user@domainname.com"
  };

  const groupsResponse = {
    value: [
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "resourceProvisioningOptions": [
          "Team"
        ]
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_MEMBER_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if required options are not passed', async () => {
    const actual = await command.validate({
      options: {
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both teamId and teamName options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        teamName: 'Team Name',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelId and channelName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelId and channelName options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        channelName: 'Channel Name',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId is invalid', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: 'Invalid',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName, userId or id options are not passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and userId options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userName: 'Demo.User@contoso.onmicrosoft.com',
        userId: '00000000-0000-0000-0000-000000000000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userName and id options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userName: 'Demo.User@contoso.onmicrosoft.com',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and id options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userId: '00000000-0000-0000-0000-000000000000',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName, userId and id options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userName: 'Demo.User@contoso.onmicrosoft.com',
        userId: '00000000-0000-0000-0000-000000000000',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userId: '00000000-0000',
        role: 'owner'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid role specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid groupId, channelId, Id and Owner role specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId, channelId, Id and Member role specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'member'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'owner'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get team when resourceprovisioning does not exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": [
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamName: 'Team Name',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      id: '00000',
      role: 'owner' } } as any), new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('correctly get teams id by team name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve(groupsResponse);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members/00000') > -1) {
        return Promise.resolve(memberResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        teamName: 'Team Name',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000',
        role: 'owner'
      }
    });
    assert(loggerLogSpy.calledWith(memberResponse));
  });

  it('fails to get channel when channel does not exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelName: 'Channel Name',
      id: '00000',
      role: 'owner' } } as any), new CommandError('The specified channel does not exist in the Microsoft Teams team'));
  });

  it('fails to get channel when channel does is not private', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('00000000-0000-0000-0000-000000000000')}/channels?$filter=displayName eq '${encodeURIComponent('Other Channel')}'`) {
        return Promise.resolve({
          "value": [
            {
              "name": "Other Channel",
              "membershipType": "standard"
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelName: 'Other Channel',
      id: '00000',
      role: 'owner' } } as any), new CommandError('The specified channel is not a private channel'));
  });

  it('correctly get channel id by channel name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "19:00000000000000000000000000000000@thread.skype",
              "membershipType": "private"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members/00000') > -1) {
        return Promise.resolve(memberResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Channel Name',
        id: '00000',
        role: 'owner'
      }
    });
    assert(loggerLogSpy.calledWith(memberResponse));
  });

  it('fails to get member when member does not exist by userId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "0",
              "displayName": "User 1",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "roles": ["owner"]
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      userId: '00000000-0000-0000-0000-000000000000',
      role: 'owner' } } as any), new CommandError('The specified member does not exist in the Microsoft Teams channel'));
  });

  it('fails to get member when member does not return userId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "0",
              "displayName": "User 1",
              "email": "user1@domainname.com",
              "roles": ["owner"]
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      userId: '00000000-0000-0000-0000-000000000000',
      role: 'owner' } } as any), new CommandError('The specified member does not exist in the Microsoft Teams channel'));
  });

  it('fails to get member when member does not exist by userName', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "0",
              "displayName": "User 1",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "roles": ["owner"]
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      userName: 'user@domainname.com',
      role: 'owner' } } as any), new CommandError('The specified member does not exist in the Microsoft Teams channel'));
  });

  it('fails to get member when member does not return email', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "0",
              "displayName": "User 1",
              "userId": "00000000-0000-0000-0000-000000000001",
              "roles": ["owner"]
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      userName: 'user@domainname.com',
      role: 'owner' } } as any), new CommandError('The specified member does not exist in the Microsoft Teams channel'));
  });

  it('fails to get member when member does multiple exist with username', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "0",
              "displayName": "User 1",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user@domainname.com"
            },
            {
              "id": "1",
              "displayName": "User 2",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user@domainname.com"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:00000000000000000000000000000000@thread.skype',
      userName: 'user@domainname.com',
      role: 'owner' } } as any), new CommandError('Multiple Microsoft Teams channel members with name user@domainname.com found: 00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'));
  });

  it('correctly get member id by user id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "00000",
              "displayName": "User",
              "userId": "00000000-0000-0000-0000-000000000000",
              "email": "user@domainname.com"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members/00000') > -1) {
        return Promise.resolve(memberResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userId: '00000000-0000-0000-0000-000000000000',
        role: 'owner'
      }
    });
    assert(loggerLogSpy.calledWith(memberResponse));
  });

  it('correctly get member id by user name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "00000",
              "displayName": "User",
              "userId": "00000000-0000-0000-0000-000000000000",
              "email": "user@domainname.com"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:00000000000000000000000000000000@thread.skype/members/00000') > -1) {
        return Promise.resolve(memberResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        userName: 'user@domainname.com',
        role: 'owner'
      }
    });
    assert(loggerLogSpy.calledWith(memberResponse));
  });

  it('correctly handles error when retrieving all teams', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: {
      debug: false,
      teamId: '00000000-0000-0000-0000-000000000000' } } as any), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
