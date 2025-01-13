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
import command from './user-guest-add.js';

describe(commands.USER_GUEST_ADD, () => {
  const emailAddress = 'john.doe@contoso.com';
  const displayName = 'John Doe';

  const requestResponse = {
    id: '7b602cb4-ccd4-40c1-a965-cc0ebaae16fd',
    inviteRedeemUrl: 'https://login.microsoftonline.com/redeem',
    invitedUserDisplayName: displayName,
    invitedUserType: 'Guest',
    invitedUserEmailAddress: emailAddress,
    sendInvitationMessage: true,
    inviteRedirectUrl: 'https://myapplications.microsoft.com',
    status: 'PendingAcceptance',
    invitedUserMessageInfo: {
      messageLanguage: 'en-US',
      customizedMessageBody: 'Could you accept this invite please?',
      ccRecipients: [
        {
          emailAddress: {
            address: emailAddress
          }
        }
      ]
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.USER_GUEST_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly logs the API response', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/invitations') {
        return requestResponse;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        emailAddress: emailAddress,
        displayName: displayName
      }
    });

    assert(loggerLogSpy.calledWith(requestResponse));
  });

  it('invites user with all options', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/invitations') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    const redirectUrl = 'https://microsoft.com';
    const welcomeMessage = 'Hello could you accept this request?';
    const ccRecipient = 'Maria.Jones@contoso.com';
    const languageCode = 'nl-BE';
    await command.action(logger, {
      options: {
        emailAddress: emailAddress,
        displayName: displayName,
        inviteRedirectUrl: redirectUrl,
        welcomeMessage: welcomeMessage,
        ccRecipients: ccRecipient,
        messageLanguage: languageCode,
        sendInvitationMessage: true
      }
    });

    const requestBody = {
      invitedUserEmailAddress: emailAddress,
      inviteRedirectUrl: redirectUrl,
      invitedUserDisplayName: displayName,
      sendInvitationMessage: true,
      invitedUserMessageInfo: {
        customizedMessageBody: welcomeMessage,
        messageLanguage: languageCode,
        ccRecipients: [{ emailAddress: { address: ccRecipient } }]
      }
    };

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('invites user with default values', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/invitations') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        emailAddress: emailAddress
      }
    });

    assert.strictEqual(postRequestStub.lastCall.args[0].data.inviteRedirectUrl, 'https://myapplications.microsoft.com');
    assert.strictEqual(postRequestStub.lastCall.args[0].data.invitedUserMessageInfo.messageLanguage, 'en-US');
  });

  it('invites user with ccRecipients', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/invitations') {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    const ccRecipient = 'Maria.Jones@contoso.com';
    await command.action(logger, {
      options: {
        emailAddress: emailAddress,
        ccRecipients: ccRecipient
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data.invitedUserMessageInfo.ccRecipients, [{ emailAddress: { address: ccRecipient } }]);
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects({ error: { message: errorMessage } });

    await assert.rejects(command.action(logger, {
      options: {
        emailAddress: emailAddress
      }
    }), new CommandError(errorMessage));
  });
});
