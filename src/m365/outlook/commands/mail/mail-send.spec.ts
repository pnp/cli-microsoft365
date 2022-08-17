import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./mail-send');

describe(commands.MAIL_SEND, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MAIL_SEND), true);
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(commands.SENDMAIL) > -1), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sends email using the basic properties', (done) => {
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

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends email using the basic properties (debug)', (done) => {
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

    command.action(logger, { options: { debug: true, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends email to multiple addresses', (done) => {
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

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t store email in sent items', (done) => {
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

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'false' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
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

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if bodyContentType is invalid', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if saveToSentItems is invalid', async () => {
    const actual = await command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'Invalid' } }, commandInfo);
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
});