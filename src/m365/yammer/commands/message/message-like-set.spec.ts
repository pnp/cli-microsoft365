import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

const command: Command = require('./message-like-set');
describe(commands.YAMMER_MESSAGE_LIKE_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let promptOptions: any;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      request.delete,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_MESSAGE_LIKE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123 } });
    assert.strictEqual(actual, true);
  });

  it('id must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('enable must be true or false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123, enable: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('enable must be true or false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123, enable: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('enable must be true or false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123, enable: 'fals' } });
    assert.notStrictEqual(actual, true);
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

  it('prompts when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { debug: false, id: 1231231, enable: 'false' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when liking a message', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1231231 } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when liking a message and confirm passed', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1231231, confirm: 'true' } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when liking a message and enabled set to true', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1231231, enable: 'true' } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when disliking a message and confirming', (done) => {
    const requestPostedStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1231231, enable: 'false', confirm: true } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts when disliking and confirmation parameter is denied', (done) => {
    cmdInstance.action({ options: { debug: false, id: 1231231, enable: 'false', confirm: false } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when disliking a message and confirmation is hit', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: true, id: 1231231, enable: 'false' } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Aborts execution when enabled set to false and confirmation is not given', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, id: 1231231, enable: 'false' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 