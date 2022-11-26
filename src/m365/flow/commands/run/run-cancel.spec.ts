import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./run-cancel');

describe(commands.RUN_CANCEL, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.RUN_CANCEL), true);
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

  it('prompts before cancelling the specified Microsoft FlowName when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        debug: false,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts cancelling the specified Microsoft FlowName when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        debug: false,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert(postSpy.notCalled);
  });

  it('cancels the specified Microsoft Flow without confirmation prompt', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/runs/08585981115186985105550762687CU161/cancel?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    });
    assert(loggerLogSpy.called);
  });

  it('cancels the specified Microsoft FlowName when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72/runs/08585981115186985105550762687CU161/cancel?api-version=2016-11-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    });
    assert(loggerLogSpy.called);
  });

  it('correctly handles no environmentName found without prompting when confirm specified', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions."
        }
      });
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    } as any), new CommandError(`You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions.`));
  });

  it('correctly handles no environmentName found when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions."
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762687CU161'
      }
    } as any), new CommandError(`You are not permitted to make flows in this 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'. Please switch to the default environment, or to one of your own environment(s), where you have maker permissions.`));
  });

  it('correctly handles specified Microsoft FlowName not found when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'."
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));


    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161'
      }
    } as any), new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'.`));
  });

  it('correctly handles specified Microsoft FlowName not found without prompting when confirm specified', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'."
        }
      });
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88',
        name: '08585981115186985105550762687CU161',
        confirm: true
      }
    } as any), new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac88' under Api 'shared_logicflows'.`));
  });

  it('correctly handles specified Microsoft FlowName run not found when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "AzureResourceManagerRequestFailed",
          "message": `Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`
        }
      });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762688CP233'
      }
    } as any), new CommandError(`Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`));
  });

  it('correctly handles specified Microsoft FlowName run not found without prompting when confirm specified', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "AzureResourceManagerRequestFailed",
          "message": `Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`
        }
      });
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: false,
        environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6',
        flowName: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        name: '08585981115186985105550762688CP233',
        confirm: true
      }
    } as any), new CommandError(`Request to Azure Resource Manager failed with error: '{"error":{"code":"WorkflowRunNotFound","message":"The workflow '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' run '08585981115186985105550762688CP233' could not be found."}}`));
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