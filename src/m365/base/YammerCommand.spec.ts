import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../telemetry.js';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import YammerCommand from './YammerCommand.js';

class MockCommand extends YammerCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }

  public handlePromiseError(response: any): void {
    this.handleRejectedODataJsonPromise(response);
  }
}

describe('YammerCommand', () => {
  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
  });

  afterEach(() => {
    sinonUtil.restore(auth.restoreAuth);
  });

  after(() => {
    sinon.restore();
  });

  it('correctly reports an error while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('doesn\'t execute command when not logged in', async () => {
    sinon.stub(auth, 'restoreAuth').resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.service.connected = false;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('executes command when logged in', async () => {
    sinon.stub(auth, 'restoreAuth').resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.service.connected = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await command.action(logger, { options: {} });
    assert(commandCommandActionSpy.called);
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://www.yammer.com/api');
  });

  it('displays error message coming from Yammer', () => {
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError({
      error: {
        base: 'abc'
      }
    }), new CommandError('abc'));
  });

  it('displays 404 error message from Yammer', () => {
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError({
      statusCode: 404
    }), new CommandError("Not found (404)"));
  });

  it('displays error message not from Yammer (1)', () => {
    const error = { error: 'not from Yammer' };
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError(error),
      new CommandError(error as any));
  });

  it('displays error message not from Yammer (2)', () => {
    const error = { message: "test" };
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError(error),
      new CommandError(error as any));
  });
});
