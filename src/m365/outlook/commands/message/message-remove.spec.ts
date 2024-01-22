import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { accessToken } from '../../../../utils/accessToken.js';
import command from './message-remove.js';

describe(commands.MESSAGE_REMOVE, () => {
  const messageId = 'AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAADb58MCAAA=';
  const userId = '6799fd1a-723b-4eb7-8e52-41ae530274ca';
  const userPrincipalName = 'john.doe@contoso.com';  

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      accessToken.isAppOnlyAccessToken,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes specific message using delegated permissions without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes specific message using delegated permissions while prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: messageId, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes specific message using delegated permissions from a shared mailbox specified by userId without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, userId: userId, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes specific message using delegated permissions from a shared mailbox specified by userPrincipalName without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, userName: userPrincipalName, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes specific message using application permissions from a mailbox specified by userId without prompting for confirmation', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, userId: userId, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes specific message using application permissions from a mailbox specified by userPrincipalName without prompting for confirmation', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/messages/${messageId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: messageId, userName: userPrincipalName, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('throws an error when both userId and userName are not defined when removing a message using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: messageId } } as any),
      new CommandError(`The option 'userId' or 'userName' is required when removing a message using application permissions`));
  });

  it('throws an error when both userId and userName are defined when removing a message using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: messageId, userId: userId, userName: userPrincipalName } } as any),
      new CommandError(`Both options 'userId' and 'userName' cannot be set when removing a message using application permissions`));
  });

  it('throws an error when userId is not a valid GUID when removing a message using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: messageId, userId: 'foo' } } as any),
      new CommandError(`The value 'foo' for 'userId' option is not a valid GUID`));
  });

  it('throws an error when userName is not a valid user principal name when removing a message using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { id: messageId, userName: 'foo' } } as any),
      new CommandError(`The value 'foo' for 'userName' option is not a valid user principal name`));
  });

  it('throws an error when both userId and userName are defined when removing a message using delegated permissions', async () => {
    await assert.rejects(command.action(logger, { options: { id: messageId, userId: userId, userName: userPrincipalName } } as any),
      new CommandError(`Both options 'userId' and 'userName' cannot be set when removing a message using delegated permissions`));
  });

  it('throws an error when userId is not a valid GUID when removing a message using delegated permissions', async () => {
    await assert.rejects(command.action(logger, { options: { id: messageId, userId: 'foo' } } as any),
      new CommandError(`The value 'foo' for 'userId' option is not a valid GUID`));
  });

  it('throws an error when userName is not a valid user principal name when removing a message using delegated permissions', async () => {
    await assert.rejects(command.action(logger, { options: { id: messageId, userName: 'foo' } } as any),
      new CommandError(`The value 'foo' for 'userName' option is not a valid user principal name`));
  });

  it('throws an error when the message cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `The specified object was not found in the store., The process failed to get the correct properties.`,
        innerError: {
          date: '2023-10-27T12:24:36',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/messages/${messageId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: messageId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the message when confirm option not passed', async () => {
    await command.action(logger, { options: { id: messageId } });

    assert(promptIssued);
  });

  it('aborts removing the message when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: messageId } });
    assert(deleteSpy.notCalled);
  });
});