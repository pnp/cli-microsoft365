import { telemetry } from '../telemetry';
import { pid } from './pid';
import { session } from './session';
import auth from '../Auth';
import sinon = require('sinon');
import { Logger } from '../cli/Logger';
import { Cli } from '../cli/Cli';
import { sinonUtil } from './sinonUtil';
import { spo } from './spo';

let cli: Cli;
export let log: string[] = [];
export let loggerLogSpy: sinon.SinonSpy;
export let loggerLogToStderrSpy: sinon.SinonSpy;
export let promptOptions: any;
export const logger: Logger = {
  log: (msg: string) => log.push(msg),
  logToStderr: (msg: string) => log.push(msg)
};

export function centralizedBeforeHook(includeRequestDigest: boolean = false): void {
  cli = Cli.getInstance();
  sinon.stub(auth, 'restoreAuth').resolves();
  sinon.stub(auth, 'storeConnectionInfo').resolves();
  sinon.stub(telemetry, 'trackEvent').returns();
  sinon.stub(pid, 'getProcessName').returns('');
  sinon.stub(session, 'getId').returns('');
  loggerLogSpy = sinon.spy(logger, 'log');
  loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  auth.service.connected = true;

  if (includeRequestDigest) {
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    sinon.stub(spo, 'ensureFormDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
  }
}

export function centralizedBeforeEachHook(includePrompt: boolean = false): void {
  log = [];
  sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  auth.service.connected = true;

  if (includePrompt) {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  }
}

export function centralizedAfterEachHook(): void {
  sinonUtil.restore([cli.getSettingWithDefaultValue]);
}

export function centralizedAfterHook(): void {
  sinon.restore();
  auth.service.connected = false;
  auth.service.spoUrl = undefined;
  auth.service.tenantId = undefined;
  auth.service.accessTokens = {};
}