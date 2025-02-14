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
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import command from './team-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TEAM_LIST, () => {
  const userId = '2630257f-11d4-4244-b4a1-3707b79f142d';
  const userName = 'john.doe@contoso.com';

  //#region mock http responses
  const groupIdResponse = {
    value: [
      { id: 'dbeea07d-d766-43b3-b6c8-0ccc22e36803' },
      { id: 'f3c7e9f3-443c-4ebc-8e8f-3fcf0b1f9b1a' },
      { id: '0d788de6-8832-4061-bfd3-1834de48967d' }
    ]
  };

  const batchResponse = {
    responses: [
      {
        id: '0',
        status: 200,
        headers: {
          'OData-Version': '4.0',
          'Content-Type': 'application/json;odata.metadata=none;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'
        },
        body: {
          id: groupIdResponse.value[0],
          createdDateTime: '2020-09-07T07:32:51.001Z',
          displayName: 'My Team',
          description: 'My Team description',
          internalId: '19:42bca71a4d84496a95b1364ed462c2eb@thread.tacv2',
          classification: null,
          specialization: 'none',
          visibility: 'private',
          webUrl: `https://teams.microsoft.com/l/team/19%3a42bca71a4d84496a95b1364ed462c2eb%40thread.tacv2/conversations?groupId=${groupIdResponse.value[0]}&tenantId=076d38cd-55b0-462f-960f-e11a369a651c`,
          isArchived: false,
          isMembershipLimitedToOwners: false,
          discoverySettings: {
            showInTeamsSearchAndSuggestions: false
          },
          memberSettings: {
            allowCreateUpdateChannels: true,
            allowCreatePrivateChannels: true,
            allowDeleteChannels: true,
            allowAddRemoveApps: true,
            allowCreateUpdateRemoveTabs: true,
            allowCreateUpdateRemoveConnectors: true
          },
          guestSettings: {
            allowCreateUpdateChannels: false,
            allowDeleteChannels: false
          },
          messagingSettings: {
            allowUserEditMessages: true,
            allowUserDeleteMessages: true,
            allowOwnerDeleteMessages: true,
            allowTeamMentions: true,
            allowChannelMentions: true
          },
          funSettings: {
            allowGiphy: true,
            giphyContentRating: 'moderate',
            allowStickersAndMemes: true,
            allowCustomMemes: true
          },
          summary: {
            ownersCount: 2,
            membersCount: 2,
            guestsCount: 0
          }
        }
      },
      {
        id: '1',
        status: 200,
        headers: {
          'OData-Version': '4.0',
          'Content-Type': 'application/json;odata.metadata=none;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'
        },
        body: {
          id: groupIdResponse.value[1],
          createdDateTime: '2020-09-07T07:32:51.001Z',
          displayName: 'My Team 3',
          description: 'My Team 3 description',
          internalId: '19:42bca71a4d84496a95b1364ed462c2eb@thread.tacv2',
          classification: null,
          specialization: 'none',
          visibility: 'private',
          webUrl: `https://teams.microsoft.com/l/team/19%3a42bca71a4d84496a95b1364ed462c2eb%40thread.tacv2/conversations?groupId=${groupIdResponse.value[1]}&tenantId=076d38cd-55b0-462f-960f-e11a369a651c`,
          isArchived: false,
          isMembershipLimitedToOwners: false,
          discoverySettings: {
            showInTeamsSearchAndSuggestions: false
          },
          memberSettings: {
            allowCreateUpdateChannels: true,
            allowCreatePrivateChannels: true,
            allowDeleteChannels: true,
            allowAddRemoveApps: true,
            allowCreateUpdateRemoveTabs: true,
            allowCreateUpdateRemoveConnectors: true
          },
          guestSettings: {
            allowCreateUpdateChannels: false,
            allowDeleteChannels: false
          },
          messagingSettings: {
            allowUserEditMessages: true,
            allowUserDeleteMessages: true,
            allowOwnerDeleteMessages: true,
            allowTeamMentions: true,
            allowChannelMentions: true
          },
          funSettings: {
            allowGiphy: true,
            giphyContentRating: 'moderate',
            allowStickersAndMemes: true,
            allowCustomMemes: true
          },
          summary: {
            ownersCount: 2,
            membersCount: 2,
            guestsCount: 0
          }
        }
      },
      {
        id: '2',
        status: 200,
        headers: {
          'OData-Version': '4.0',
          'Content-Type': 'application/json;odata.metadata=none;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'
        },
        body: {
          id: groupIdResponse.value[2],
          createdDateTime: '2020-09-07T07:32:51.001Z',
          displayName: 'My Team 2',
          description: 'My Team 2 description',
          internalId: '19:42bca71a4d84496a95b1364ed462c2eb@thread.tacv2',
          classification: null,
          specialization: 'none',
          visibility: 'private',
          webUrl: `https://teams.microsoft.com/l/team/19%3a42bca71a4d84496a95b1364ed462c2eb%40thread.tacv2/conversations?groupId=${groupIdResponse.value[2]}&tenantId=076d38cd-55b0-462f-960f-e11a369a651c`,
          isArchived: false,
          isMembershipLimitedToOwners: false,
          discoverySettings: {
            showInTeamsSearchAndSuggestions: false
          },
          memberSettings: {
            allowCreateUpdateChannels: true,
            allowCreatePrivateChannels: true,
            allowDeleteChannels: true,
            allowAddRemoveApps: true,
            allowCreateUpdateRemoveTabs: true,
            allowCreateUpdateRemoveConnectors: true
          },
          guestSettings: {
            allowCreateUpdateChannels: false,
            allowDeleteChannels: false
          },
          messagingSettings: {
            allowUserEditMessages: true,
            allowUserDeleteMessages: true,
            allowOwnerDeleteMessages: true,
            allowTeamMentions: true,
            allowChannelMentions: true
          },
          funSettings: {
            allowGiphy: true,
            giphyContentRating: 'moderate',
            allowStickersAndMemes: true,
            allowCustomMemes: true
          },
          summary: {
            ownersCount: 2,
            membersCount: 2,
            guestsCount: 0
          }
        }
      }
    ]
  };

  const commandResponse = batchResponse.responses.map(r => r.body).sort((x: any, y: any) => x.displayName.localeCompare(y.displayName));
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let postStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').returnsArg(1);
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
    (command as any).items = [];

    postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return batchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isArchived', 'description']);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { joined: true, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { joined: true, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation userName is used without joined or associated option', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation userId is used without joined or associated option', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and userName are used', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { joined: true, userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both joined and associated options are used', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { joined: true, associated: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is used with joined option', async () => {
    const actual = await command.validate({ options: { joined: true, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userName is used with associated option', async () => {
    const actual = await command.validate({ options: { associated: true, userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only joined option is used', async () => {
    const actual = await command.validate({ options: { joined: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only associated option is used', async () => {
    const actual = await command.validate({ options: { associated: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves all teams in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id&$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`) {
        return groupIdResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: groupIdResponse.value.map((obj, index) => ({
        id: index.toString(),
        method: 'GET',
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        url: `/teams/${obj.id}`
      }))
    });
    assert(loggerLogSpy.calledOnceWith(commandResponse));
  });

  it('retrieves all joined teams for the current logged in user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id`) {
        return groupIdResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        joined: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: groupIdResponse.value.map((obj, index) => ({
        id: index.toString(),
        method: 'GET',
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        url: `/teams/${obj.id}`
      }))
    });
    assert(loggerLogSpy.calledOnceWith(commandResponse));
  });

  it('retrieves all joined teams for a specified user by UPN', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/joinedTeams?$select=id`) {
        return groupIdResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userName: userName,
        joined: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: groupIdResponse.value.map((obj, index) => ({
        id: index.toString(),
        method: 'GET',
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        url: `/teams/${obj.id}`
      }))
    });
    assert(loggerLogSpy.calledOnceWith(commandResponse));
  });

  it('retrieves all associated teams for the current logged in user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/teamwork/associatedTeams?$select=id`) {
        return groupIdResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        associated: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: groupIdResponse.value.map((obj, index) => ({
        id: index.toString(),
        method: 'GET',
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        url: `/teams/${obj.id}`
      }))
    });
    assert(loggerLogSpy.calledOnceWith(commandResponse));
  });

  it('retrieves all associated teams for a specified user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/teamwork/associatedTeams?$select=id`) {
        return groupIdResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        userId: userId,
        associated: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: groupIdResponse.value.map((obj, index) => ({
        id: index.toString(),
        method: 'GET',
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        url: `/teams/${obj.id}`
      }))
    });
    assert(loggerLogSpy.calledOnceWith(commandResponse));
  });

  it('handles API error correctly', async () => {
    postStub.restore();

    const failedResponse: any = { ...batchResponse };
    failedResponse.responses[1] = {
      id: '0',
      status: 404,
      headers: {
        'OData-Version': '4.0',
        'Content-Type': 'application/json;odata.metadata=none;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'
      },
      body: {
        error: {
          message: 'Resource not found.'
        }
      }
    };

    sinon.stub(request, 'post').resolves(failedResponse);
    sinon.stub(request, 'get').resolves(groupIdResponse);

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Resource not found.'));
  });
});
