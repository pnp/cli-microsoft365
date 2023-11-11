import assert from 'assert';
import sinon from 'sinon';
import auth, { CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import PowerAppsCommand from './PowerAppsCommand.js';

class MockCommand extends PowerAppsCommand {
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

describe('PowerAppsCommand', () => {
  const cmd = new MockCommand();
  const cloudError = new CommandError(`Power Apps commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
  });

  after(() => {
    sinon.restore();
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.powerapps.com');
  });

  it(`doesn't throw error when not connected`, () => {
    auth.service.active = false;
    (cmd as any).initAction({ options: {} }, {});
  });

  it('throws error when connected to USGov cloud', () => {
    auth.service.active = true;
    auth.service.cloudType = CloudType.USGov;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovHigh cloud', () => {
    auth.service.active = true;
    auth.service.cloudType = CloudType.USGovHigh;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to USGovDoD cloud', () => {
    auth.service.active = true;
    auth.service.cloudType = CloudType.USGovDoD;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it('throws error when connected to China cloud', () => {
    auth.service.active = true;
    auth.service.cloudType = CloudType.China;
    assert.throws(() => (cmd as any).initAction({ options: {} }, {}), cloudError);
  });

  it(`doesn't throw error when connected to public cloud`, () => {
    auth.service.active = true;
    auth.service.cloudType = CloudType.Public;
    assert.doesNotThrow(() => (cmd as any).initAction({ options: {} }, {}));
  });
});
