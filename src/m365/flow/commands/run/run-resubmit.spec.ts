import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./run-resubmit');

describe(commands.RUN_RESUBMIT, () => {
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
    sinonUtil.restore([
      request.post,
      request.get,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.RUN_RESUBMIT), true);
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

  it('prompts before resubmitting the specified Microsoft Flow when confirm option not passed', (done) => {
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

  it('aborts resubmitting the specified Microsoft Flow when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'post');
    const getSpy = sinon.spy(request, 'get');
    sinonUtil.restore(Cli.prompt);
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
        assert(getSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found when prompt confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions."
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles specified Microsoft Flow not found when prompt confirmed', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'."
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles specified Microsoft Flow run not found when prompt confirmed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "name": "manual",
                "id": "/providers/Microsoft.ProcessSimple/environments//Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual",
                "type": "Microsoft.ProcessSimple/environments/flows/triggers",
                "properties": {
                  "provisioningState": "Succeeded",
                  "createdTime": "2020-10-23T23:16:15.131033Z",
                  "changedTime": "2020-10-23T23:22:13.3611905Z",
                  "state": "Enabled"
                }
              }
            ]
          });
        }
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "AzureResourceManagerRequestFailed",
          "message": `Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: false,
        environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762688CP233'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly getting triggername for the specified Microsoft Flow when prompt confirmed (debug)', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "name": "manual",
                "id": "/providers/Microsoft.ProcessSimple/environments//Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual",
                "type": "Microsoft.ProcessSimple/environments/flows/triggers",
                "properties": {
                  "provisioningState": "Succeeded",
                  "createdTime": "2020-10-23T23:16:15.131033Z",
                  "changedTime": "2020-10-23T23:22:13.3611905Z",
                  "state": "Enabled"
                }
              }
            ]
          });
        }
      }
      return Promise.reject('Invalid request');
    });

    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 202 });
      }
      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options:
      {
        debug: true,
        environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161'
      }
    }, () => {
      try {
        assert.notStrictEqual(loggerLogToStderrSpy.getCall(1).args[0].indexOf('Retrieved trigger: manual'), -1);
        assert(getStub.called);
        assert(postStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('resubmits the specified Microsoft Flow run when confirm specified (debug)', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "name": "manual",
                "id": "/providers/Microsoft.ProcessSimple/environments//Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual",
                "type": "Microsoft.ProcessSimple/environments/flows/triggers",
                "properties": {
                  "provisioningState": "Succeeded",
                  "createdTime": "2020-10-23T23:16:15.131033Z",
                  "changedTime": "2020-10-23T23:22:13.3611905Z",
                  "state": "Enabled"
                }
              }
            ]
          });
        }
      }
      return Promise.reject('Invalid request');
    });

    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 202 });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options:
      {
        debug: true,
        environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flow: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01');
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01');
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