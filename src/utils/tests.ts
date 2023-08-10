import { telemetry } from '../telemetry';
import { pid } from './pid';
import { session } from './session';
import auth from '../Auth';
import sinon = require('sinon');
import { Logger } from '../cli/Logger';
import { Cli } from '../cli/Cli';
import { sinonUtil } from './sinonUtil';

let cli: Cli;
export let log: string[] = [];
export let loggerLogSpy: sinon.SinonSpy;
export let loggerLogToStderrSpy: sinon.SinonSpy;
export const logger: Logger = {
  log: (msg: string) => log.push(msg),
  logToStderr: (msg: string) => log.push(msg)
};

export function includeDefaultBeforeHookSetup(): void {
  cli = Cli.getInstance();
  sinon.stub(auth, 'restoreAuth').resolves();
  sinon.stub(auth, 'storeConnectionInfo').resolves();
  sinon.stub(telemetry, 'trackEvent').returns();
  sinon.stub(pid, 'getProcessName').returns('');
  sinon.stub(session, 'getId').returns('');
  loggerLogSpy = sinon.spy(logger, 'log');
  loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  auth.service.connected = true;
}

export function includeDefaultBeforeEachHookSetup(): void {
  log = [];
  sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  auth.service.connected = true;
}

export function includeDefaultAfterEachHookSetup(): void {
  sinonUtil.restore([cli.getSettingWithDefaultValue]);
}

export function includeDefaultAfterHookSetup(): void {
  sinon.restore();
  auth.service.connected = false;
  auth.service.spoUrl = undefined;
  auth.service.tenantId = undefined;
  auth.service.accessTokens = {};
}