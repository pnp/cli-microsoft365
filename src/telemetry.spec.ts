import * as sinon from 'sinon';
import * as assert from 'assert';
import appInsights from './appInsights';
import { Cli } from "./cli/Cli";
import { settingsNames } from './settingsNames';
import { sinonUtil } from './utils/sinonUtil';
import { telemetry } from './telemetry';

describe('Telemetry', () => {
  afterEach(() => {
    sinonUtil.restore([
      Cli.getInstance().getSettingWithDefaultValue,
      appInsights.trackEvent,
      appInsights.trackException
    ]);
  });

  it(`doesn't log an event when disableTelemetry is set`, async () => {
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    const trackEventStub = sinon.stub(appInsights, 'trackEvent').callsFake(() => { });

    telemetry.trackEvent('foo bar', {});

    assert(trackEventStub.notCalled);
  });

  it('logs an event when disableTelemetry is not set', async () => {
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    const trackEventStub = sinon.stub(appInsights, 'trackEvent').callsFake(() => { });

    telemetry.trackEvent('foo bar', {});

    assert(trackEventStub.called);
  });

  it(`doesn't log an exception when disableTelemetry is set`, async () => {
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    const exceptionStub = sinon.stub(appInsights, 'trackException').callsFake(() => { });

    telemetry.trackException('Error!');

    assert(exceptionStub.notCalled);
  });

  it('logs an exception when disableTelemetry is not set', async () => {
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    const trackExceptionStub = sinon.stub(appInsights, 'trackException').callsFake(() => { });

    telemetry.trackException('Error!');

    assert(trackExceptionStub.called);
  });
});