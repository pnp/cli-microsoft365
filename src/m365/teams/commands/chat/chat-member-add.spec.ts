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
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./chat-member-add');

describe(commands.CHAT_MEMBER_ADD, () => {
  const chatId = '19:09fd7575940146d383a4a83fc9598546@thread.v2';
  const userPrincipalName = 'john@contoso.com';
  const userId = 'a857e888-b602-4790-86d9-3dca2109449e';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHAT_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a member by specifying its userId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['guest'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
      visibleHistoryStartDateTime: undefined
    };

    await command.action(logger, { options: { chatId: chatId, userId: userId, role: 'guest', verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('adds a member by specifying its userName', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userPrincipalName)}`,
      visibleHistoryStartDateTime: undefined
    };

    await command.action(logger, { options: { chatId: chatId, userName: userPrincipalName, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('adds a member by specifying its userId and share all chat history', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
      visibleHistoryStartDateTime: '0001-01-01T00:00:00Z'
    };

    await command.action(logger, { options: { chatId: chatId, userId: userId, includeAllHistory: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('adds a member by specifying its userId with chat history from a certain date', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
      visibleHistoryStartDateTime: '2019-04-18T23:51:43.255Z'
    };

    await command.action(logger, { options: { chatId: chatId, userId: userId, visibleHistoryStartDateTime: '2019-04-18T23:51:43.255Z' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('fails validation if chatId is not valid chatId', async () => {
    const actual = await command.validate({ options: { chatId: 'invalid', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not valid guid', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not valid UPN', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userName: userPrincipalName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if role is not a valid role', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId, role: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if role is a valid role', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId, role: 'guest' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the visibleHistoryStartDateTime is not a valid date', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId, visibleHistoryStartDateTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the visibleHistoryStartDateTime is a valid date', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId, visibleHistoryStartDateTime: '2019-04-18T23:51:43.255Z' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles OData API error', async () => {
    const error = {
      error: {
        code: 'generalException',
        message: 'The member can\'t be added to the chat'
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(error.error.message));
  });
});
