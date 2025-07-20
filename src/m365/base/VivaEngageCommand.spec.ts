import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../telemetry.js';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import VivaEngageCommand from './VivaEngageCommand.js';
import { accessToken } from '../../utils/accessToken.js';

class MockCommand extends VivaEngageCommand {
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

describe('VivaEngageCommand', () => {
  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  afterEach(() => {
    sinonUtil.restore(auth.restoreAuth);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
    auth.connection.active = false;
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
    auth.connection.active = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await command.action(logger, { options: {} });
    assert(commandCommandActionSpy.called);
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://www.yammer.com/api');
  });

  it('displays error message coming from Viva Engage', () => {
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError({
      error: {
        base: 'abc'
      }
    }), new CommandError('abc'));
  });

  it('displays 404 error message from Viva Engage', () => {
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError({
      statusCode: 404
    }), new CommandError("Not found (404)"));
  });

  it('displays error message not from Viva Engage (1)', () => {
    const error = { error: 'not from Viva Engage' };
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError(error),
      new CommandError(error as any));
  });

  it('displays error message not from Viva Engage (2)', () => {
    const error = { message: "test" };
    const mock = new MockCommand();
    assert.throws(() => mock.handlePromiseError(error),
      new CommandError(error as any));
  });

  it('throws error when using application-only permissions', async () => {
    const cmd = new MockCommand();
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    auth.connection.active = true;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), new CommandError('This command requires delegated permissions.'));
  });
});
