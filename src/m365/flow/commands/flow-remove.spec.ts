import commands from '../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./flow-remove');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';
import * as chalk from 'chalk';

describe(commands.FLOW_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    promptOptions = undefined;
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name and environment specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Flow owned by the currently signed-in user when confirm option not passed', (done) => {
    cmdInstance.action({
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
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

  it('aborts removing the specified Microsoft Flow owned by the currently signed-in user when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
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

  it('removes the specified Microsoft Flow owned by the currently signed-in user when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        debug: true,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified Microsoft Flow owned by another user when confirm option not passed', (done) => {
    cmdInstance.action({
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
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

  it('aborts removing the specified Microsoft Flow owned by another user when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
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

  it('removes the specified Microsoft Flow owned by another user when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        debug: true,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified Microsoft Flow without prompting when confirm specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    }, () => {
      assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
      done();
    }, (err: any) => done(err));
  });

  it('removes the specified Microsoft Flow as Admin without prompting when confirm specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true,
        asAdmin: true
      }
    }, () => {
      assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
      done();
    }, (err: any) => done(err));
  });

  it('correctly handles no environment found without prompting when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
        }
      });
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
        }
      });
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Microsoft Flow found when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.resolve({ statusCode: 204 });
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.red(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no Microsoft Flow found when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.resolve({ statusCode: 204 });
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.red(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});