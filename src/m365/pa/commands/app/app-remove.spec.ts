import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', async () => {
    const actual = await command.validate({
      options: {
        name: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Power App when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified Microsoft Power App when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified Microsoft Power App when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
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
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Power App from other user when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
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
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Power App without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft PowerApp from other user without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2017-08-01`) {
        return Promise.resolve({ statusCode: 200 });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles no Microsoft Power App found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({ response: { status: 403 } });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any), new CommandError(`App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`));
  });

  it('correctly handles no Microsoft Power App found when confirm specified', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({ response: { status: 403 } });
    });

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any), new CommandError(`App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`));
  });

  it('correctly handles Microsoft Power App found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 200 });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any);
  });

  it('correctly handles Microsoft Power App found when confirm specified', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.resolve({ statusCode: 200 });
    });

    await command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        confirm: true
      }
    } as any);
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

  it('supports specifying confirm', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('correctly handles random api error', async () => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject("Something went wrong");
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any), new CommandError("Something went wrong"));
  });
});
