import assert from 'assert';
import sinon from 'sinon';
import auth, { CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import PowerAutomateCommand from './PowerAutomateCommand.js';
import { accessToken } from '../../utils/accessToken.js';
import { sinonUtil } from '../../utils/sinonUtil.js';

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

describe('PowerAutomateCommand', () => {
  const cmd = new MockCommand();
  const cloudError = new CommandError(`Power Automate commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
    auth.service.connected = true;
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.flow.microsoft.com');
  });

  it(`doesn't throw error when not connected`, () => {
    auth.connection.active = false;
    (cmd as any).initAction({ options: {} }, {});
  });

  it('throws error when connected to USGov cloud', () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGov;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovHigh cloud', () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGovHigh;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovDoD cloud', () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGovDoD;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to China cloud', () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.China;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it(`doesn't throw error when connected to public cloud`, () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.Public;
    assert.doesNotThrow(() => (cmd as any).initAction({ options: {} }, {}));
  });

  it('throws error when trying to use the command using application only permissions', () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    auth.service.connected = true;
    auth.service.cloudType = CloudType.Public;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}));
  });
});
