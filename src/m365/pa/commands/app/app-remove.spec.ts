import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-remove.js';

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
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
    assert.strictEqual(command.name, commands.APP_REMOVE);
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
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, {
      options: {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the specified Microsoft Power App when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Power App from other user when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Power App from other user as admin when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/admin/environments/4ce50206-9576-4237-8b17-38d8aadfaa35/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        environmentName: '4ce50206-9576-4237-8b17-38d8aadfaa35',
        asAdmin: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft Power App without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/e0c89645-7f00-4877-a290-cbaf6e060da1?api-version=2017-08-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes the specified Microsoft PowerApp from other user without prompting when confirm specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72?api-version=2017-08-01`) {
        return { statusCode: 200 };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: '0f64d9dd-01bb-4c1b-95b3-cb4a1a08ac72',
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles no Microsoft Power App found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').rejects({ response: { status: 403 } });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any), new CommandError(`App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`));
  });

  it('correctly handles no Microsoft Power App found when confirm specified', async () => {
    sinon.stub(request, 'delete').rejects({ response: { status: 403 } });

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        force: true
      }
    } as any), new CommandError(`App 'e0c89645-7f00-4877-a290-cbaf6e060da1' does not exist`));
  });

  it('correctly handles Microsoft Power App found when prompt confirmed', async () => {
    sinon.stub(request, 'delete').resolves({ statusCode: 200 });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any);
  });

  it('correctly handles Microsoft Power App found when confirm specified', async () => {
    sinon.stub(request, 'delete').resolves({ statusCode: 200 });

    await command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1',
        force: true
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
      if (o.option.indexOf('--force') > -1) {
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

  it('supports specifying asAdmin', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--asAdmin') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('correctly handles random api error', async () => {
    sinon.stub(request, 'delete').rejects(new Error("Something went wrong"));

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(command.action(logger, {
      options:
      {
        name: 'e0c89645-7f00-4877-a290-cbaf6e060da1'
      }
    } as any), new CommandError("Something went wrong"));
  });

  it('fails validation if asAdmin specified without environment', async () => {
    const actual = await command.validate({ options: { name: "5369f386-e380-46cb-82a4-4e18f9e4f3a7", asAdmin: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if environment specified without admin', async () => {
    const actual = await command.validate({ options: { name: "5369f386-e380-46cb-82a4-4e18f9e4f3a7", environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if asAdmin specified with environment', async () => {
    const actual = await command.validate({ options: { name: "5369f386-e380-46cb-82a4-4e18f9e4f3a7", asAdmin: true, environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
