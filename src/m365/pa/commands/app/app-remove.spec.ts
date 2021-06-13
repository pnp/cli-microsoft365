import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.delete,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', () => {
    const actual = command.validate({
      options: {
        name: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', () => {
    const actual = command.validate({
      options: {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Power App when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    }, () => {
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

  it('aborts removing the specified Microsoft Power App when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified Microsoft Power App when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified Microsoft Power App from other user when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified Microsoft Power App without prompting when confirm specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    }, () => {
      assert(loggerLogToStderrSpy.called);
      done();
    });
  });

  it('removes the specified Microsoft PowerApp from other user without prompting when confirm specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    }, () => {
      assert(loggerLogToStderrSpy.called);
      done();
    });
  });

  it('correctly handles no Microsoft Power App found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({ response: { status: 403 } });
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err, `App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Microsoft Power App found when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({ response: { status: 403 } });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err, `App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Microsoft Power App found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 200 });
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Microsoft Power App found when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 200 });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirm', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
