import assert from 'assert';
import sinon from 'sinon';
import { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './cli-reconsent.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';

describe(commands.RECONSENT, () => {
  let log: string[];
  let logger: Logger;
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
    sinonUtil.restore([
      cli.getTenant,
      cli.getClientId
    ]);
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
    sinon.stub(cli, 'getClientId').returns('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(cli, 'getTenant').returns('common');
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(`To re-consent your Microsoft Entra application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent`));
  });

  it('shows message with url (using autoOpenLinksInBrowser)', async () => {
    sinon.stub(cli, 'getClientId').returns('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(cli, 'getTenant').returns('common');
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
    sinon.stub(cli, 'getClientId').returns('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(cli, 'getTenant').returns('common');
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
