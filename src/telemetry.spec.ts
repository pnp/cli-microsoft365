import assert from 'assert';
import sinon from 'sinon';
import appInsights from './appInsights.js';
import { cli } from "./cli/cli.js";
import { settingsNames } from './settingsNames.js';
import { telemetry } from './telemetry.js';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';
import { sinonUtil } from './utils/sinonUtil.js';

describe('Telemetry', () => {
  let trackEventStub: sinon.SinonStub;
  let trackExceptionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('abc123');
    trackEventStub = sinon.stub(appInsights, 'trackEvent');
    trackExceptionStub = sinon.stub(appInsights, 'trackException');
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.getSettingWithDefaultValue,
      (telemetry as any).trackTelemetry
    ]);
    trackEventStub.resetHistory();
    trackExceptionStub.resetHistory();
  });

  after(() => {
    sinon.restore();
  });

  it(`doesn't log an event when disableTelemetry is set`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    await telemetry.trackEvent('foo bar', {});
    assert(trackEventStub.notCalled, 'trackEventStub called');
    assert(trackExceptionStub.notCalled, 'trackExceptionStub called');
  });

  it('logs an event when disableTelemetry is not set', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    await telemetry.trackEvent('foo bar', {});
    assert(trackEventStub.called, 'trackEventStub not called');
    assert(trackExceptionStub.notCalled, 'trackExceptionStub called');
  });

  it(`doesn't log the exception when disableTelemetry is set`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    await telemetry.trackEvent('foo bar', {}, 'error');
    assert(trackEventStub.notCalled, 'trackEventStub called');
    assert(trackExceptionStub.notCalled, 'trackExceptionStub called');
  });

  it('logs the exception when disableTelemetry is not set', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    await telemetry.trackEvent('foo bar', {}, 'error');
    assert(trackEventStub.notCalled, 'trackEventStub called');
    assert(trackExceptionStub.called, 'trackExceptionStub not called');
  });

  it(`logs an empty string for shell if it couldn't resolve shell process name`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    sinonUtil.restore(pid.getProcessName);
    sinon.stub(pid, 'getProcessName').returns(undefined);

    await telemetry.trackEvent('foo bar', {});
    assert.strictEqual(appInsights.commonProperties.shell, '');
  });

  it(`fails silently when submitting telemetry fails`, async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    sinonUtil.restore(appInsights.trackEvent);
    sinon.stub(appInsights, 'trackEvent').throws(new Error('foo'));

    await telemetry.trackEvent('foo bar', {});
    assert.ok(true);
  });
});