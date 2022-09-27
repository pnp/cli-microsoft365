import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./completion-sh-setup');

describe(commands.COMPLETION_SH_SETUP, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let generateShCompletionStub: sinon.SinonStub;
  let setupShCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    generateShCompletionStub = sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
    setupShCompletionStub = sinon.stub(autocomplete, 'setupShCompletion').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    generateShCompletionStub.reset();
    setupShCompletionStub.reset();
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      autocomplete.generateShCompletion,
      autocomplete.setupShCompletion
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_SH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('generates file with commands info', async () => {
    await command.action(logger, { options: { debug: false } });
    assert(generateShCompletionStub.called);
  });

  it('sets up command completion in the shell', async () => {
    await command.action(logger, { options: { debug: false } });
    assert(setupShCompletionStub.called);
  });

  it('writes additional info in debug mode', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.calledWith('Generating command completion...'));
  });
});