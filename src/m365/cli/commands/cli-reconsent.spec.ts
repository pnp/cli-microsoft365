import assert from 'assert';
import sinon from 'sinon';
import { CommandError } from '../../../Command.js';
import { Cli } from '../../../cli/Cli.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './cli-reconsent.js';

describe(commands.RECONSENT, () => {
  let log: string[];
  let logger: Logger;
  let cli: Cli;
  let getSettingWithDefaultValueStub: sinon.SinonStub;
  let loggerLogSpy: sinon.SinonSpy;
  let openStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
  });

  beforeEach(() => {
    log = [];
    cli = Cli.getInstance();
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => false));
    openStub = sinon.stub(browserUtil, 'open').callsFake(async () => { return; });
  });

  afterEach(() => {
    loggerLogSpy.restore();
    getSettingWithDefaultValueStub.restore();
    openStub.restore();
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.RECONSENT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows message with url (not using autoOpenLinksInBrowser)', async () => {
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(`To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent`));
  });

  it('shows message with url (using autoOpenLinksInBrowser)', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => true));

    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === 'https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent') {
        return;
      }
      throw 'Invalid url';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent`));
  });

  it('throws error when open in browser fails', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => true));

    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === 'https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent') {
        throw 'An error occurred';
      }
      throw 'Invalid url';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error occurred'));
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent`));
  });
});
