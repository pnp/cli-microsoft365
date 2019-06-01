import * as sinon from 'sinon';
import * as assert from 'assert';
import AadCommand from './AadCommand';
import auth from './AadAuth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';
import appInsights from '../../appInsights';

class MockCommand extends AadCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('AadCommand', () => {
  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
  });

  after(() => {
    Utils.restore(appInsights.trackEvent);
  });
  
  it('correctly reports an error while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'aad command'
      },
      log: (msg: any) => {},
      prompt: () => {},
      action: command.action()
    };
    cmdInstance.action({options:{}}, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.restoreAuth);
      }
    });
  });

  it('doesn\'t execute command when error occurred while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const cmdInstance = {
      commandWrapper: {
        command: 'aad command'
      },
      log: (msg: any) => {},
      prompt: () => {},
      action: command.action()
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    cmdInstance.action({options:{}}, () => {
      try {
        assert(commandCommandActionSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(auth.restoreAuth);
      }
    });
  });
});