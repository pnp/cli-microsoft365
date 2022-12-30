import * as assert from 'assert';
import * as sinon from 'sinon';
import auth, { CloudType } from '../../Auth';
import { CommandError } from '../../Command';
import { telemetry } from '../../telemetry';
import { sinonUtil } from '../../utils/sinonUtil';
import PowerPlatformCommand from './PowerPlatformCommand';

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
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent
    ]);
  });

  it('returns correct bapi resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.bap.microsoft.com');
  });

  it(`doesn't throw error when not connected`, () => {
    auth.service.connected = false;
    (cmd as any).initAction({ options: {} }, {});
  });

  it('throws error when connected to USGov cloud', () => {
    auth.service.connected = true;
    auth.service.cloudType = CloudType.USGov;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovHigh cloud', () => {
    auth.service.connected = true;
    auth.service.cloudType = CloudType.USGovHigh;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovDoD cloud', () => {
    auth.service.connected = true;
    auth.service.cloudType = CloudType.USGovDoD;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to China cloud', () => {
    auth.service.connected = true;
    auth.service.cloudType = CloudType.China;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it(`doesn't throw error when connected to public cloud`, () => {
    auth.service.connected = true;
    auth.service.cloudType = CloudType.Public;
    assert.doesNotThrow(() => (cmd as any).initAction({ options: {} }, {}));
  });
});
