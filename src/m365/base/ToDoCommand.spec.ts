import assert from 'assert';
import sinon from 'sinon';
import auth from '../../Auth.js';
import { telemetry } from '../../telemetry.js';
import PowerAutomateCommand from './PowerAutomateCommand.js';
import { accessToken } from '../../utils/accessToken.js';

class MockCommand extends PowerAutomateCommand {
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
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
    auth.connection.active = true;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('throws error when trying to use the command using application only permissions', () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    auth.connection.active = true;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}));
  });
});
