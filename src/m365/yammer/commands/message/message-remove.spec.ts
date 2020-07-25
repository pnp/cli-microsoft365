import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_MESSAGE_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;

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
        cb({ continue: false });
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.delete
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_MESSAGE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('id must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('id is required', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('calls the messaging endpoint with the right parameters and confirmation', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 10123190123123, confirm: true } }, () => {
      try {
        assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the messaging endpoint with the right parameters without confirmation', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: true, id: 10123190123123, confirm: false } }, () => {
      try {
        assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('does not call the messaging endpoint without confirmation', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };

    cmdInstance.action({ options: { debug: true, id: 10123190123123, confirm: false } }, () => {
      try {
        assert(requestDeleteStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, id: 10123190123123, confirm: true } }, (err?: any) => {
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