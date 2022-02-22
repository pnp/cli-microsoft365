import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Logger } from '../../../cli';
import Utils from '../../../Utils';
import PowerPlatformCommand from './PowerPlatformCommand';

class MockCommand extends PowerPlatformCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    cb();
  }

  public commandHelp(): void {
  }
}

describe('PowerPlatformCommand', () => {
  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  afterEach(() => {
    Utils.restore(auth.restoreAuth);
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  });

  it('doesn\'t execute command when not logged in', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
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
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
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

  it('returns correct bapi resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.bap.microsoft.com');
  });
});
