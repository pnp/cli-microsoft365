import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-guest-add');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_GUEST_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'inviteRedeemUrl', 'invitedUserDisplayName', 'invitedUserEmailAddress', 'invitedUserType', 'resetRedemption', 'sendInvitationMessage', 'status']);
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
    sinon.stub(request, 'post').callsFake(async () => { throw { error: { message: errorMessage } }; });

    await assert.rejects(command.action(logger, {
      options: {
        emailAddress: emailAddress
      }
    }), new CommandError(errorMessage));
  });
});