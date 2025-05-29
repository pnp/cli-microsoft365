import assert from 'assert';
import sinon from 'sinon';
import auth, { CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import PowerPlatformCommand from './PowerPlatformCommand.js';
import { accessToken } from '../../utils/accessToken.js';
import { sinonUtil } from '../../utils/sinonUtil.js';

class MockCommand extends PowerPlatformCommand {
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

describe('PowerPlatformCommand', () => {
  const cmd = new MockCommand();
  const cloudError = new CommandError(`Power Platform commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  beforeEach(() => {
    auth.connection.active = true;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('returns correct bapi resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.bap.microsoft.com');
  });

  it(`doesn't throw error when not connected`, async () => {
    auth.connection.active = false;
    await (cmd as any).initAction({ options: {} }, {});
  });

  it('throws error when connected to USGov cloud', async () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGov;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovHigh cloud', async () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGovHigh;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovDoD cloud', async () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.USGovDoD;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to China cloud', async () => {
    auth.connection.active = true;
    auth.connection.cloudType = CloudType.China;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it(`doesn't throw error when connected to public cloud`, () => {
    auth.connection.cloudType = CloudType.Public;
    assert.doesNotThrow(() => (cmd as any).initAction({ options: {} }, {}));
  });

  it('throws error when using application-only permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    auth.connection.cloudType = CloudType.Public;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), new CommandError('This command requires delegated permissions.'));
  });
});
