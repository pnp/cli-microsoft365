import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./mail-send');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.OUTLOOK_MAIL_SEND, () => {
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      fs.readFileSync,
      fs.existsSync,
      fs.lstatSync,
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OUTLOOK_MAIL_SEND), true);
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(commands.OUTLOOK_SENDMAIL) > -1), true);
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
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
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
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends email using contents from a file', (done) => {
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
    sinon.stub(fs, 'readFileSync').callsFake(() => 'Lorem ipsum');
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContentsFilePath: 'file.txt' } }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sends HTML email using contents from a file', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      message: {
        subject: 'Lorem ipsum',
        body: {
          contentType: 'HTML',
          content: 'Lorem <b>ipsum</b>'
        },
        toRecipients: [{ emailAddress: { address: 'mail@domain.com' } }]
      },
      saveToSentItems: undefined
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'Lorem ipsum');
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem <b>ipsum</b>', bodyContentType: 'HTML' } }, () => {
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
    sinon.stub(fs, 'readFileSync').callsFake(() => 'Lorem ipsum');
    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bodyContents: 'Lorem ipsum' } }, () => {
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
      actual = JSON.stringify(opts.body);
      if (opts.url === `https://graph.microsoft.com/v1.0/me/sendMail`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'false' } }, () => {
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
    sinon.stub(request, 'post').callsFake((opts) => {
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

    cmdInstance.action({ options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither bodyContents nor bodyContentsFilePath specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if bodyContents and bodyContentsFilePath specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentsFilePath: 'file.txt' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified bodyContentsFilePath doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContentsFilePath: 'file.txt' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified bodyContentsFilePath points to a folder', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContentsFilePath: 'file' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if bodyContentType is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if saveToSentItems is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when subject, to and bodyContents are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple to emails are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com,mail2@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple to emails separated with command and space are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com, mail2@domain.com', bodyContents: 'Lorem ipsum' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the specified bodyContentsFilePath points to a file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com, mail2@domain.com', bodyContentsFilePath: 'file.txt' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when bodyContentType is set to Text', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Text' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when bodyContentType is set to HTML', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'HTML' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when saveToSentItems is set to false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when saveToSentItems is set to true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', saveToSentItems: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});