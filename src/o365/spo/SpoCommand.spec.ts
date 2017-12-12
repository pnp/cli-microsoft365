import * as sinon from 'sinon';
import * as assert from 'assert';
import SpoCommand from './SpoCommand';
import auth from './SpoAuth';
import Utils from '../../Utils';
import { CommandError } from '../../Command';

class MockCommand extends SpoCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
  }
}

describe('SpoCommand', () => {
  it('correctly reports an error while restoring auth info', (done) => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const cmdInstance = {
      log: (msg: any) => {},
      prompt: () => {},
      action: command.action()
    };
    const cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    cmdInstance.action({options:{}}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
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