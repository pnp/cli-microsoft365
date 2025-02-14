import assert from 'assert';
import sinon from 'sinon';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import commands from '../../commands.js';
import command from './completion-sh-setup.js';

describe(commands.COMPLETION_SH_SETUP, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let generateShCompletionStub: sinon.SinonStub;
  let setupShCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    generateShCompletionStub = sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
    setupShCompletionStub = sinon.stub(autocomplete, 'setupShCompletion').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    generateShCompletionStub.reset();
    setupShCompletionStub.reset();
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_SH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('generates file with commands info', async () => {
    await command.action(logger, { options: {} });
    assert(generateShCompletionStub.called);
  });

  it('sets up command completion in the shell', async () => {
    await command.action(logger, { options: {} });
    assert(setupShCompletionStub.called);
  });

  it('writes additional info in debug mode', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.calledWith('Generating command completion...'));
  });
});
