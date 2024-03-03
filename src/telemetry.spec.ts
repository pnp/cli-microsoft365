import assert from 'assert';
import child_process from 'child_process';
import sinon from 'sinon';
import { cli } from "./cli/cli.js";
import { settingsNames } from './settingsNames.js';
import { telemetry } from './telemetry.js';
import { pid } from './utils/pid.js';
import { sinonUtil } from './utils/sinonUtil.js';
import { session } from './utils/session.js';

describe('Telemetry', () => {
  let spawnStub: sinon.SinonStub;
  let stdin: string = '';

  before(() => {
    spawnStub = sinon.stub(child_process, 'spawn').callsFake(() => {
      return {
        stdin: {
          write: (s: string) => {
            stdin += s;
          },
          end: () => { }
        },
        unref: () => { }
      } as any;
    });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => 'abc123');
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.getSettingWithDefaultValue,
      (telemetry as any).trackTelemetry
    ]);
    spawnStub.resetHistory();
    stdin = '';
  });

  after(() => {
    sinon.restore();
  });

  it(`doesn't log an event when disableTelemetry is set`, () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    telemetry.trackEvent('foo bar', {});
    assert(spawnStub.notCalled);
  });

  it('logs an event when disableTelemetry is not set', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    telemetry.trackEvent('foo bar', {});
    assert(spawnStub.called);
  });

  it(`logs an empty string for shell if it couldn't resolve shell process name`, () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    sinonUtil.restore(pid.getProcessName);
    sinon.stub(pid, 'getProcessName').callsFake(() => undefined);

    telemetry.trackEvent('foo bar', {});
    assert.strictEqual(JSON.parse(stdin).shell, '');
  });

  it(`silently handles exception if an error occurs while spawning telemetry runner`, (done) => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    sinonUtil.restore(child_process.spawn);
    sinon.stub(child_process, 'spawn').throws();
    try {
      telemetry.trackEvent('foo bar', {});
      done();
    }
    catch (e) {
      done(e);
    }
  });
});