import assert from 'assert';
import sinon from 'sinon';
import auth from '../../Auth.js';
import { telemetry } from '../../telemetry.js';
import DelegatedGraphCommand from './DelegatedGraphCommand.js';
import { accessToken } from '../../utils/accessToken.js';
import { CommandError } from '../../Command.js';

class MockCommand extends DelegatedGraphCommand {
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

describe('ToDoCommand', () => {
  const cmd = new MockCommand();

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it(`doesn't throw error when not connected`, () => {
    auth.connection.active = false;
    (cmd as any).initAction({ options: {} }, {});
    auth.connection.active = true;
  });

  it('throws error when using application-only permissions', () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), new CommandError('This command does not support application-only permissions.'));
  });
});
