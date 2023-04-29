import Command from '../../Command';
import commands from './commands';
import * as assert from 'assert';
import * as sinon from 'sinon';
import { Logger } from '../../cli/Logger';
import * as open from 'open';
import { telemetry } from '../../telemetry';
import { pid } from '../../utils/pid';
import { session } from '../../utils/session';
import { sinonUtil } from '../../utils/sinonUtil';
import { Cli } from '../../cli/Cli';
const packageJSON = require('../../../package.json');
const command: Command = require('./docs');

describe(commands.DOCS, () => {
  let log: any[];
  let logger: Logger;
  let cli: Cli;
  let loggerLogSpy: sinon.SinonSpy;
  let getSettingWithDefaultValueStub: sinon.SinonStub;
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
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any)._open = open;
    openStub = sinon.stub(command as any, '_open').callsFake(() => Promise.resolve(null));
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').returns(false);
  });

  afterEach(() => {
    loggerLogSpy.restore();
    getSettingWithDefaultValueStub.restore();
    openStub.restore();
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DOCS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should log a message and return if autoOpenLinksInBrowser is false', async () => {
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(`Use a web browser to open the CLI for Microsoft 365 docs webpage URL`));
  });

  it('should open the CLI for Microsoft 365 docs webpage URL using "open" if autoOpenLinksInBrowser is true', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').returns(true);
    await command.action(logger, { options: {} });
    assert(openStub.calledWith(packageJSON.homepage), 'open should have been called with the CLI for Microsoft 365 docs webpage URL');
  });
});