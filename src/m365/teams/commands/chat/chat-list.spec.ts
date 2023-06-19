import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./chat-list');

describe(commands.CHAT_LIST, () => {
  const userId = '63be605f-94c6-433b-b763-22bb16dd4acf';
  const userName = 'user@contoso.com';
  const chatsResponse = [
    {
      "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2",
      "topic": "Meeting chat sample",
      "createdDateTime": "2020-12-08T23:53:05.801Z",
      "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z",
      "chatType": "meeting"
    },
    {
      "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2",
      "topic": "Group chat sample",
      "createdDateTime": "2020-12-03T19:41:07.054Z",
      "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z",
      "chatType": "group"
    },
    {
      "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces",
      "topic": null,
      "createdDateTime": "2020-12-04T23:10:28.51Z",
      "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z",
      "chatType": "oneOnOne"
    }
  ];
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHAT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'topic', 'chatType']);
  });

  it('fails validation for an incorrect chatType.', async () => {
    const actual = await command.validate({ options: { type: 'oneOn' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId and userName are specified', async () => {
    const actual = await command.validate({ options: { userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input without chat type', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with a userId defined', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for oneOnOne chat conversations with a specific userName defined', async () => {
    const actual = await command.validate({ options: { type: "oneOnOne", userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for group chat conversation', async () => {
    const actual = await command.validate({ options: { type: "group" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for meeting chat conversations', async () => {
    const actual = await command.validate({ options: { type: "meeting" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all chat conversations for the currently signed in user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/chats`) {
        return { 'value': chatsResponse };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(chatsResponse));
  });

  it('lists oneOnOne chat conversations for the currently signed in user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/chats?$filter=chatType eq 'oneOnOne'`) {
        return { 'value': chatsResponse.filter(y => y.chatType === 'oneOnOne') };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: { type: 'oneOnOne' }
    });
    assert(loggerLogSpy.calledWith(chatsResponse.filter(y => y.chatType === 'oneOnOne')));
  });

  it('lists group chat conversations for the currently signed in user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/chats?$filter=chatType eq 'group'`) {
        return { 'value': chatsResponse.filter(y => y.chatType === 'group') };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { type: 'group' } });
    assert(loggerLogSpy.calledWith(chatsResponse.filter(y => y.chatType === 'group')));
  });

  it('lists group chat conversations for a specific user by id', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/chats?$filter=chatType eq 'group'`) {
        return { 'value': chatsResponse.filter(y => y.chatType === 'group') };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { type: 'group', userId: userId } });
    assert(loggerLogSpy.calledWith(chatsResponse.filter(y => y.chatType === 'group')));
  });

  it('lists meeting chat conversations for a specific user by userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/chats?$filter=chatType eq 'meeting'`) {
        return { 'value': chatsResponse.filter(y => y.chatType === 'meeting') };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { type: 'meeting', userName: userName } });
    assert(loggerLogSpy.calledWith(chatsResponse.filter(y => y.chatType === 'meeting')));
  });


  it('outputs all data in json output mode for the currently signed in user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/chats`) {
        return {
          'value': chatsResponse
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith(chatsResponse));
  });

  it('correctly handles error when listing chat conversations', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('throws an error when passing userId using delegated permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    await assert.rejects(command.action(logger, { options: { userId: userId } } as any), new CommandError(`The options 'userId' or 'userName' cannot be used when obtaining chats using delegated permissions`));
  });

  it('throws an error when not passing userId or userName using application permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(`The option 'userId' or 'userName' is required when obtaining chats using app only permissions`));
  });
});
