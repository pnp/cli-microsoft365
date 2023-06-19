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
import { odata } from '../../../../utils/odata';
const command: Command = require('./chat-member-remove');

describe(commands.CHAT_MEMBER_REMOVE, () => {
  const chatId = '19:09fd7575940146d383a4a83fc9598546@thread.v2';
  const userPrincipalName = 'john@contoso.com';
  const userId = 'a857e888-b602-4790-86d9-3dca2109449e';
  const chatMemberId = 'MCMjMCMjZTFkZDQwMjMtYTY1Ni00ODBhLThhMGUtYzFiMWVlYzUxZTFkIyMxOTowOWZkNzU3NTk0MDE0NmQzODNhNGE4M2ZjOTU5ODU0NkB0aHJlYWQudjIjI2E4NTdlODg4LWI2MDItNDc5MC04NmQ5LTNkY2EyMTA5NDQ5ZQ=="';
  const chatMembers = [
    {
      id: chatMemberId,
      roles: ['owner'],
      displayName: 'John Doe',
      visibleHistoryStartDateTime: '2022-04-08T09:15:53.423Z',
      userId: userId,
      email: userPrincipalName,
      tenantId: 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d'
    },
    {
      id: 'MCMjMCMjZTFkZDQwMjMtYTY1Ni00ODBhLThhMGUtYzFiMWVlYzUxZTFkIyMxOTowOWZkNzU3NTk0MDE0NmQzODNhNGE4M2ZjOTU5ODU0NkB0aHJlYWQudjIjI2ZlMzZmNzVlLWMxMDMtNDEwYi1hMThhLTJiZjZkZjA2YWMzYQ==',
      roles: ['owner'],
      displayName: 'Adele Vance',
      visibleHistoryStartDateTime: '2022-04-08T09:15:53.423Z',
      userId: 'fe36f75e-c103-410b-a18a-2bf6df06ac3a',
      email: 'adele@contoso.com',
      tenantId: 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d'
    }
  ];

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

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
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHAT_MEMBER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the member by specifying the userId', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return chatMembers;
      }

      throw 'Invalid request';
    });
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members/${chatMemberId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { chatId: chatId, userId: userId, confirm: true, verbose: true } });
    assert(deleteStub.called);
  });

  it('removes the member from a chat by specifying the member id', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members/${chatMemberId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { chatId: chatId, id: chatMemberId, confirm: true, verbose: true } });
    assert(deleteStub.called);
  });


  it('removes the specified member retrieved by user principal name when prompt confirmed', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return chatMembers;
      }

      throw 'Invalid request';
    });
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members/${chatMemberId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { chatId: chatId, userName: userPrincipalName, verbose: true } });
    assert(deleteStub.called);
  });

  it('throws error when member by specifying userName is not found in the chat', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return [...chatMembers.slice(1)];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { chatId: chatId, userName: userPrincipalName, confirm: true, verbose: true } }),
      new CommandError(`Member with userName '${userPrincipalName}' could not be found in the chat.`));
  });

  it('throws error when member by specifying userId is not found in the chat', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/chats/${chatId}/members`) {
        return [...chatMembers.slice(1)];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { chatId: chatId, userId: userId, confirm: true, verbose: true } }),
      new CommandError(`Member with userId '${userId}' could not be found in the chat.`));
  });

  it('prompts before removing the specified message when confirm option not passed', async () => {
    await command.action(logger, { options: { chatId: chatId, id: chatMemberId } });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified chat member when confirm option not passed and prompt not confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { chatId: chatId, userId: userId } });
    assert(deleteStub.notCalled);
  });

  it('fails validation if the chatId is not valid chatId', async () => {
    const actual = await command.validate({ options: { chatId: 'invalid', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userId is not valid guid', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not valid UPN', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if ID of a chat member is passed', async () => {
    const actual = await command.validate({ options: { chatId: chatId, id: chatMemberId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { chatId: chatId, userName: userPrincipalName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
