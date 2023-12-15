import assert from 'assert';
import sinon from 'sinon';
import auth, { CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import PowerAutomateCommand from './PowerAutomateCommand.js';

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
  });

  after(() => {
    sinon.restore();
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.flow.microsoft.com');
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
