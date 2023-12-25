import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './run-resubmit.js';

describe(commands.RUN_RESUBMIT, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RUN_RESUBMIT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the flowName is not valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: 'invalid',
        name: '08585981115186985105550762687CU161'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name, environmentName and flowName specified', async () => {
    const actual = await command.validate({
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resubmitting the specified Microsoft Flow when force option not passed', async () => {
    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });


    assert(promptIssued);
  });

  it('aborts resubmitting the specified Microsoft Flow when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    const getSpy = sinon.spy(request, 'get');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert(postSpy.notCalled);
    assert(getSpy.notCalled);
  });

  it('correctly handles no environment found when prompt confirmed', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions."
      }
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    } as any), new CommandError(`You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions.`));
  });

  it('correctly handles specified Microsoft Flow not found when prompt confirmed', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "ConnectionAuthorizationFailed",
        "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'."
      }
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161'
      }
    } as any), new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'.`));
  });

  it('correctly handles specified Microsoft Flow run not found when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').rejects({
      "error": {
        "code": "AzureResourceManagerRequestFailed",
        "message": `Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`
      }
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762688CP233'
      }
    } as any), new CommandError(`Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`));
  });

  it('correctly getting triggername for the specified Microsoft Flow when prompt confirmed (debug)', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01`) {
        return { statusCode: 202 };
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options:
      {
        debug: true,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert.notStrictEqual(loggerLogToStderrSpy.getCall(1).args[0].indexOf('Retrieved trigger: manual'), -1);
    assert(getStub.called);
    assert(postStub.called);
  });

  it('resubmits the specified Microsoft Flow run when confirm specified (debug)', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
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
          };
        }
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01`) {
        return { statusCode: 202 };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        debug: true,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161',
        force: true
      }
    });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers?api-version=2016-11-01');
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c6/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88/triggers/manual/histories/08585981115186985105550762687CU161/resubmit?api-version=2016-11-01');
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying flow', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--flow') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
