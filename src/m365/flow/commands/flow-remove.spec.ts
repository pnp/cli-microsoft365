import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../telemetry';
import auth from '../../../Auth';
import { Cli } from '../../../cli/Cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import { pid } from '../../../utils/pid';
import { session } from '../../../utils/session';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
const command: Command = require('./flow-remove');

describe(commands.REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name and environment specified', async () => {
    const actual = await command.validate({
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Flow owned by the currently signed-in user when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified Microsoft Flow owned by the currently signed-in user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified Microsoft Flow owned by the currently signed-in user when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('prompts before removing the specified Microsoft Flow owned by another user when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified Microsoft Flow owned by another user when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, {
      options: {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified Microsoft Flow owned by another user when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        asAdmin: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Flow without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Flow as Admin without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c/flows/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2016-11-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true,
        asAdmin: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles no environment found without prompting when confirm specified', async () => {
    sinon.stub(request, 'delete').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
      }
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    } as any), new CommandError(`Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied.`));
  });

  it('correctly handles no environment found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied."
      }
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    } as any), new CommandError(`Access to the environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c' is denied.`));
  });

  it('correctly handles no Microsoft Flow found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').resolves({ statusCode: 204 });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72'
      }
    } as any), new CommandError(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`));
  });

  it('correctly handles no Microsoft Flow found when confirm specified', async () => {
    sinon.stub(request, 'delete').resolves({ statusCode: 204 });

    await assert.rejects(command.action(logger, {
      options:
      {
        environmentName: 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c',
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    } as any), new CommandError(`Error: Resource '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72' does not exist in environment 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c'`));
  });
});
