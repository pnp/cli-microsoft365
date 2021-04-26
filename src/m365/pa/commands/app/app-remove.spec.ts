import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
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
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name and environment specified', () => {
    const actual = command.validate({
      options: {
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Power App when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
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
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
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
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/Default-8063a435-fc8f-447b-b03b-9e50a265c748/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2016-11-01`) {
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
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
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
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/Default-8063a435-fc8f-447b-b03b-9e50a265c748/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    }, () => {
      assert(loggerLogToStderrSpy.called);
      done();
    });
  });

  it('correctly handles no environment found without prompting when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748' is denied."
        }
      });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748' is denied."
        }
      });
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Microsoft Power App found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 204 });
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(chalk.red(`Error: Resource 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist in environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Microsoft Power App found when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 204 });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-8063a435-fc8f-447b-b03b-9e50a265c748',
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(chalk.red(`Error: Resource 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist in environment 'Default-8063a435-fc8f-447b-b03b-9e50a265c748'`)));
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

  it('supports specifying environment', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});