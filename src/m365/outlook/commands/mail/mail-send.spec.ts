import * as assert from 'assert';
import * as sinon from 'sinon';
import * as fs from 'fs';
import appInsights from '../../../../appInsights';
import auth, { Auth } from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./mail-send');

describe(commands.MAIL_SEND, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
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
    (command as any).items = [];
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Auth.isAppOnlyAuth,
      fs.existsSync,
      fs.readFileSync,
      fs.lstatSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MAIL_SEND), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sends email using the basic properties', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email using the basic properties (debug)', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email to multiple addresses', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [
          { emailAddress: { address: 'mail@domain.com' } },
          { emailAddress: { address: 'mail2@domain.com' } }
        ]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email to multiple cc recipients', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [
          { emailAddress: { address: 'mail@domain.com' } },
          { emailAddress: { address: 'mail2@domain.com' } }
        ],
        ccRecipients: [
          { emailAddress: { address: 'mail3@domain.com' } },
          { emailAddress: { address: 'mail4@domain.com' } }
        ]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', cc: 'mail3@domain.com,mail4@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email to multiple bcc recipients', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [
          { emailAddress: { address: 'mail@domain.com' } },
          { emailAddress: { address: 'mail2@domain.com' } }
        ],
        bccRecipients: [
          { emailAddress: { address: 'mail3@domain.com' } },
          { emailAddress: { address: 'mail4@domain.com' } }
        ]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bcc: 'mail3@domain.com,mail4@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('doesn\'t store email in sent items', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: 'false'
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'false' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email with multiple attachments', async () => {
    const fileContentBase64 = 'TG9yZW0gaXBzdW0gZG9sb3Igc2l0IGFtZXQsIGNvbnNlY3RldHVyIGFkaXBpc2NpbmcgZWxpdC4=';
    sinon.stub(fs, 'readFileSync').returns(fileContentBase64);

    const requestPostStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        subject: 'Lorem ipsum',
        to: 'mail@domain.com',
        bodyContents: 'Lorem ipsum',
        attachment: ['C:/File1.txt', 'C:/File2.txt']
      }
    });
    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data.message.attachments, [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'File1.txt', contentBytes: fileContentBase64 }, { '@odata.type': '#microsoft.graph.fileAttachment', name: 'File2.txt', contentBytes: fileContentBase64 }]);
  });

  it('sends email with single attachment', async () => {
    const fileContentBase64 = 'TG9yZW0gaXBzdW0gZG9sb3Igc2l0IGFtZXQsIGNvbnNlY3RldHVyIGFkaXBpc2NpbmcgZWxpdC4=';
    sinon.stub(fs, 'readFileSync').returns(fileContentBase64);

    const requestPostStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        subject: 'Lorem ipsum',
        to: 'mail@domain.com',
        bodyContents: 'Lorem ipsum',
        attachment: 'C:/File1.txt'
      }
    });
    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data.message.attachments, [{ '@odata.type': '#microsoft.graph.fileAttachment', name: 'File1.txt', contentBytes: fileContentBase64 }]);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } } as any),
      new CommandError(`An error has occurred`));
  });

  it('fails validation if bodyContentType is invalid', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if saveToSentItems is invalid', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if importance is invalid', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', importance: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if file doesn\'t exist', async () => {
    sinon.stub(fs, 'lstatSync').returns({ isFile: () => true } as any);
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString() === 'C:/File2.txt') {
        return false;
      }

      return true;
    });

    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', attachment: ['C:/File.txt', 'C:/File2.txt'] } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails vlaidation if attachment is not a file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').callsFake(path => {
      if (path.toString() === 'C:/File2.txt') {
        return { isFile: () => false } as any;
      }

      return { isFile: () => true } as any;
    });

    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', attachment: ['C:/File.txt', 'C:/File2.txt'] } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if attachments are too large', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns({ isFile: () => true } as any);
    sinon.stub(fs, 'readFileSync').callsFake(path => {
      if (path.toString() === 'C:/File.txt') {
        return 'A'.repeat(4_250_000);
      }

      throw 'Invalid read request';
    });

    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', attachment: 'C:/File.txt' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when subject, to and bodyContents are specified', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple to emails are specified', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bodyContents: 'Lorem ipsum' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple to emails separated with command and space are specified', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com, mail2@domain.com', bodyContents: 'Lorem ipsum' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when bodyContentType is set to Text', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Text' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when bodyContentType is set to HTML', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'HTML' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when saveToSentItems is set to false', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when saveToSentItems is set to true', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
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

  it('sends email using a specified group mailbox', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        from: { emailAddress: { address: 'sales@domain.com' } },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', mailbox: 'sales@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('sends email using a specified sender', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'Text',
          content: 'Lorem ipsum'
        },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: undefined
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter('some-user@domain.com')}/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', sender: 'some-user@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, expected);
  });

  it('throws an error when the sender is not defined when signed in using app only authentication', async () => {
    sinonUtil.restore([Auth.isAppOnlyAuth]);
    sinon.stub(Auth, 'isAppOnlyAuth').callsFake(() => true);

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        subject: 'Lorem ipsum',
        to: 'mail@domain.com',
        bodyContents: 'Lorem ipsum'
      }
    } as any), new CommandError(`Specify a upn or user id in the 'sender' option when using app only authentication.`));
  });
});