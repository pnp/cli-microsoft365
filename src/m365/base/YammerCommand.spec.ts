import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli';
import { CommandError } from '../../Command';
import Utils from '../../Utils';
import YammerCommand from './YammerCommand';

class MockCommand extends YammerCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(logger: Logger, args: {}, cb: () => void): void {
    cb();
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }

  public handlePromiseError(response: any, logger: Logger, callback: (err?: any) => void): void {
    this.handleRejectedODataJsonPromise(response, logger, callback);
  }
}

describe('YammerCommand', () => {
  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  afterEach(() => {
    Utils.restore(auth.restoreAuth);
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  });

  it('correctly reports an error while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: (msg: any) => { },
      logRaw: (msg: any) => { },
      logToStderr: (msg: any) => { }
    };
    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t execute command when error occurred while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: (msg: any) => { },
      logRaw: (msg: any) => { },
      logToStderr: (msg: any) => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    command.action(logger, { options: {} }, () => {
      try {
        assert(commandCommandActionSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t execute command when not logged in', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    const command = new MockCommand();
    const logger: Logger = {
      log: (msg: any) => { },
      logRaw: (msg: any) => { },
      logToStderr: (msg: any) => { }
    };
    auth.service.connected = false;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    command.action(logger, { options: {} }, () => {
      try {
        assert(commandCommandActionSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes command when logged in', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    const command = new MockCommand();
    const logger: Logger = {
      log: (msg: any) => { },
      logRaw: (msg: any) => { },
      logToStderr: (msg: any) => { }
    };
    auth.service.connected = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    command.action(logger, { options: {} }, () => {
      try {
        assert(commandCommandActionSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://www.yammer.com/api');
  });

  it('displays error message coming from Yammer', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      logRaw: (msg?: string) => { },
      logToStderr: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand();
    mock.handlePromiseError({
      error: {
        base: 'abc'
      }
    }, cmd, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
    });

  });

  it('displays 404 error message from Yammer', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      logRaw: (msg?: string) => { },
      logToStderr: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand();
    mock.handlePromiseError({
      statusCode: 404
    }, cmd, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Not found (404)")));
    });
  });

  it('displays error message not from Yammer (1)', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      logRaw: (msg?: string) => { },
      logToStderr: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand();
    mock.handlePromiseError({
      error: 'not from Yammer'
    }, cmd, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify({ "message": { "error": "not from Yammer" } }));
    });
  });

  it('displays error message not from Yammer (2)', () => {
    const cmd = {
      commandWrapper: {
        command: 'command'
      },
      log: (msg?: string) => { },
      logRaw: (msg?: string) => { },
      logToStderr: (msg?: string) => { },
      prompt: () => { }
    };
    const mock = new MockCommand();
    mock.handlePromiseError({
      message: "test"
    }, cmd, (err?: any) => {
      assert.strictEqual(JSON.stringify(err), JSON.stringify({ "message": { "message": "test" } }));
    });
  });
});