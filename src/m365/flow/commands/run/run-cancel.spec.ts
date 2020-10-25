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
const command: Command = require('./run-cancel');

describe(commands.FLOW_RUN_CANCEL, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
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
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.post,
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_RUN_CANCEL), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the flow is not valid GUID', () => {
    const actual = command.validate({
      options: {
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: 'invalid',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the name is not valid RUN ID', () => {
    const actual = command.validate({
      options: {
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: 'invalid',
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('passes validation when the name, environment and flow specified', () => {
    const actual = command.validate({
      options: {
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('prompts before cancelling the specified Microsoft Flow when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
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

  it('aborts cancelling the specified Microsoft Flow when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'post');
    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
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

  it('cancels the specified Microsoft Flow when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/runs/08585981115186985105550762687CU161/cancel?api-version=2016-11-01`) {
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
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    }, () => {
      try {
        assert(loggerSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Cancelling the specified Microsoft Flow without prompting when confirm specified (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/runs/08585981115186985105550762687CU161/cancel?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    }, () => {
      assert(loggerSpy.calledWith(chalk.green('DONE')));
      done();
    });
  });

  it('correctly handles no environment found without prompting when confirm specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
        }
      });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    } as any, (err?: any) => {
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
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
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
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  /*Check if flow is cancelled*/
  /*it('correctly handles no Microsoft Flow found when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve({ statusCode: 403 });
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
      }
    } as any, (err?: any) => {
      try {
        assert(loggerSpy.calledWith(chalk.red(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });*/

  /*check if flow is cancelled*/
  /*it('correctly handles no Microsoft Flow found when confirm specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve({ statusCode: 403 });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    } as any, (err?: any) => {
      try {
        assert(loggerSpy.calledWith(chalk.red(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });*/

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

  it('supports specifying flow', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--flow') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});