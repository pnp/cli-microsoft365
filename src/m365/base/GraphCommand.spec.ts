import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import { sinonUtil } from '../../utils/sinonUtil';
import GraphCommand from './GraphCommand';

class MockCommand extends GraphCommand {
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
}

describe('GraphCommand', () => {
  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  afterEach(() => {
    sinonUtil.restore(auth.restoreAuth);
  });

  after(() => {
    sinonUtil.restore(appInsights.trackEvent);
  });

  it('correctly reports an error while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('doesn\'t execute command when not logged in', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    auth.service.connected = false;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('executes command when logged in', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    const command = new MockCommand();
    const logger: Logger = {
      log: () => { },
      logRaw: () => { },
      logToStderr: () => { }
    };
    auth.service.connected = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await command.action(logger, { options: {} });
    assert(commandCommandActionSpy.called);
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://graph.microsoft.com');
  });
});