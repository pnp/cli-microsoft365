import { telemetry } from '../telemetry';
import { pid } from './pid';
import { session } from './session';
import auth from '../Auth';
import sinon = require('sinon');
import { Logger } from '../cli/Logger';
import { Cli } from '../cli/Cli';
import { sinonUtil } from './sinonUtil';

export interface CentralizedTestSetup {
  log: () => string[];
  logger: Logger;
  loggerLogSpy: sinon.SinonSpy;
  loggerLogToStderrSpy: sinon.SinonSpy;
  runBeforeEachHookDefaults: () => void;
  runAfterEachHookDefaults: () => void;
  runAfterHookDefaults: () => void;
}

export function initializeTestSetup(): CentralizedTestSetup {
  const cli = Cli.getInstance();
  const logger: Logger = {
    log: (msg: string) => log.push(msg),
    logToStderr: (msg: string) => log.push(msg)
  };
  let log: string[] = [];
  const loggerLogSpy: sinon.SinonSpy = sinon.spy(logger, 'log');
  const loggerLogToStderrSpy: sinon.SinonSpy = sinon.spy(logger, 'logToStderr');
  sinon.stub(auth, 'restoreAuth').resolves();
  sinon.stub(auth, 'storeConnectionInfo').resolves();
  sinon.stub(telemetry, 'trackEvent').returns();
  sinon.stub(pid, 'getProcessName').returns('');
  sinon.stub(session, 'getId').returns('');
  auth.service.connected = true;

  return {
    logger,
    log: () => log,
    loggerLogSpy,
    loggerLogToStderrSpy,
    runBeforeEachHookDefaults: (): void => {
      log = [];
      sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
      auth.service.connected = true;
    },
    runAfterEachHookDefaults: (): void => {
      sinonUtil.restore([cli.getSettingWithDefaultValue]);
    },
    runAfterHookDefaults: (): void => {
      sinon.restore();
      auth.service.connected = false;
      auth.service.spoUrl = undefined;
      auth.service.tenantId = undefined;
      auth.service.accessTokens = {};
    }
  };
}