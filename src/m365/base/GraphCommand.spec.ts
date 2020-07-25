import * as sinon from 'sinon';
import * as assert from 'assert';
import GraphCommand from './GraphCommand';
import auth from '../../Auth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';
import appInsights from '../../appInsights';
import { CommandInstance } from '../../cli';

class MockCommand extends GraphCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    cb();
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('GraphCommand', () => {
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
    const cmdInstance = {
      commandWrapper: {
        command: 'graph command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    cmdInstance.action({ options: {} }, (err?: any) => {
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
    const cmdInstance = {
      commandWrapper: {
        command: 'graph command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    cmdInstance.action({ options: {} }, () => {
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
    const cmdInstance = {
      commandWrapper: {
        command: 'graph command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    auth.service.connected = false;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    cmdInstance.action({ options: {} }, () => {
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
    const cmdInstance = {
      commandWrapper: {
        command: 'graph command'
      },
      log: (msg: any) => { },
      prompt: () => { },
      action: command.action()
    };
    auth.service.connected = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    cmdInstance.action({ options: {} }, () => {
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
    assert.strictEqual((command as any).resource, 'https://graph.microsoft.com');
  });
});